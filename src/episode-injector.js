#!/usr/bin/env node
/**
 * episode-injector.js
 *
 * Clones named overlay Groups from an existing .tscproj and places them
 * at calculated timestamps.
 *
 * For Mac .cmproj bundles, it updates project.tscproj in place and writes
 * a project.tscproj.bak backup before overwriting.
 *
 * Usage:
 *   node episode-injector.js episode.cmproj/project.tscproj
 *
 * No dependencies. Requires Node 14+.
 */

"use strict";

const fs = require("fs");
const path = require("path");

// ---------------------------------------------------------------------------
// CONSTANTS
// ---------------------------------------------------------------------------

const TICKS_PER_SECOND = 705_600_000;
const CAMTASIA_INT64_MAX = "9223372036854775807";
const JS_ROUNDED_INT64_MAX = "9223372036854776000";

// ---------------------------------------------------------------------------
// LOGGING
// ---------------------------------------------------------------------------

function log(level, msg) {
  const prefix = {
    step: "\n>> ",
    ok: "  [ok]    ",
    info: "  [info]  ",
    warn: "  [warn]  ",
    error: "  [error] ",
    sub: "          -> ",
    detail: "             ",
  };
  console.log(`${prefix[level] ?? "  "}${msg}`);
}

// ---------------------------------------------------------------------------
// TICK HELPERS
// ---------------------------------------------------------------------------

function secondsToTicks(s) {
  return Math.round(s * TICKS_PER_SECOND);
}

function ticksToTimestamp(ticks) {
  const total = ticks / TICKS_PER_SECOND;
  const m = Math.floor(total / 60);
  const s = (total % 60).toFixed(1).padStart(4, "0");
  return `${m}:${s}`;
}

// ---------------------------------------------------------------------------
// FILE SYSTEM
// ---------------------------------------------------------------------------

/** Recursively copy a directory. */
function copyDirSync(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDirSync(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

// ---------------------------------------------------------------------------
// PROJECT HELPERS
// ---------------------------------------------------------------------------

function getSceneTracks(project) {
  return project.timeline.sceneTrack.scenes[0].csml.tracks;
}

function getTrackAttributes(project) {
  return project.timeline.trackAttributes ?? [];
}

function getVideoDuration(tracks) {
  let max = 0;
  function recurse(trackList) {
    for (const track of trackList) {
      for (const media of track.medias) {
        const end = (media.start ?? 0) + (media.duration ?? 0);
        if (end > max) max = end;
        if (media._type === "Group" && media.tracks) recurse(media.tracks);
      }
    }
  }
  recurse(tracks);
  return max;
}

function getContentDuration(tracks) {
  // Determine episode content end from host/guest proxy groups.
  // Looks for Groups whose ident contains HostName, GuestName, or EpisodeName
  // (e.g. HostNameOverlay), which typically span the full content duration.
  // Falls back to getVideoDuration if no matching groups are found.
  const patterns = [/HostName/i, /GuestName/i, /EpisodeName/i];
  let max = 0;
  for (const track of tracks) {
    for (const media of track.medias) {
      if (media._type === "Group") {
        const ident = media.attributes?.ident ?? "";
        if (patterns.some((p) => p.test(ident))) {
          const end = (media.start ?? 0) + (media.duration ?? 0);
          if (end > max) max = end;
        }
      }
    }
  }
  return max > 0 ? max : getVideoDuration(tracks);
}

function findMaxId(tracks) {
  let max = 0;
  function recurse(trackList) {
    for (const track of trackList) {
      for (const media of track.medias) {
        if (media.id > max) max = media.id;
        if (media._type === "Group" && media.tracks) recurse(media.tracks);
      }
    }
  }
  recurse(tracks);
  return max;
}

function findOverlayByIdent(tracks, ident, preferTrackIndex = null) {
  // Collect all matching Groups, then prefer the one on preferTrackIndex if given.
  const matches = [];
  for (const track of tracks) {
    for (const media of track.medias) {
      if (media._type === "Group" && media.attributes?.ident === ident) {
        matches.push({ media, track });
      }
    }
  }
  if (matches.length === 0) return null;
  if (preferTrackIndex !== null) {
    const preferred = matches.find((m) => m.track.trackIndex === preferTrackIndex);
    if (preferred) return preferred;
  }
  return matches[0];
}

function rangesOverlap(startA, durationA, startB, durationB) {
  const endA = startA + durationA;
  const endB = startB + durationB;
  return startA < endB && startB < endA;
}

function findCollisionInTrack(track, start, duration) {
  for (const media of track.medias) {
    if (rangesOverlap(start, duration, media.start ?? 0, media.duration ?? 0)) {
      return media;
    }
  }
  return null;
}

function findCollisionInTrackSet(trackSet, start, duration) {
  for (const track of trackSet) {
    const collision = findCollisionInTrack(track, start, duration);
    if (collision) {
      return { track, media: collision };
    }
  }
  return null;
}

function enforceMutexGroupNoOverlap(tracks, rules) {
  const rulesByMutexGroup = new Map();
  for (const rule of rules) {
    if (!rule.mutexGroup) continue;
    if (!rulesByMutexGroup.has(rule.mutexGroup)) {
      rulesByMutexGroup.set(rule.mutexGroup, []);
    }
    rulesByMutexGroup.get(rule.mutexGroup).push(rule);
  }

  for (const [groupName, groupRules] of rulesByMutexGroup.entries()) {
    const idents = new Set(groupRules.map((r) => r.ident));
    const trackIndexes = new Set(groupRules.map((r) => r.trackIndex));

    const medias = [];
    for (const track of tracks) {
      if (!trackIndexes.has(track.trackIndex)) continue;
      for (const media of track.medias ?? []) {
        if (media._type !== "Group") continue;
        if (!idents.has(media.attributes?.ident)) continue;
        medias.push(media);
      }
    }

    medias.sort((a, b) => (a.start ?? 0) - (b.start ?? 0));

    let moved = 0;
    let lastEnd = 0;
    for (const media of medias) {
      const start = media.start ?? 0;
      const duration = media.duration ?? 0;
      if (start < lastEnd) {
        media.start = lastEnd;
        moved++;
      }
      lastEnd = (media.start ?? 0) + duration;
    }

    if (moved > 0) {
      log("info", `Mutex group "${groupName}": shifted ${moved} media(s) to eliminate remaining overlaps`);
    }
  }
}

function alignNameOverlayToTrackLength(tracks, trackAttributes, sourceTrackName, overlayTrackName, overlayIdentRegex) {
  const sourceTrack = tracks.find(
    (t) => (trackAttributes[t.trackIndex]?.ident ?? "").toLowerCase() === sourceTrackName.toLowerCase()
  );
  if (!sourceTrack) {
    log("warn", `${sourceTrackName} track not found — cannot align ${overlayTrackName}`);
    return;
  }

  const overlayTrack = tracks.find(
    (t) => (trackAttributes[t.trackIndex]?.ident ?? "").toLowerCase() === overlayTrackName.toLowerCase()
  );
  if (!overlayTrack) {
    log("warn", `${overlayTrackName} track not found — cannot align overlay`);
    return;
  }

  let sourceEnd = 0;
  for (const media of sourceTrack.medias ?? []) {
    const end = (media.start ?? 0) + (media.duration ?? 0);
    if (end > sourceEnd) sourceEnd = end;
  }
  if (sourceEnd <= 0) {
    log("warn", `${sourceTrackName} track has no media duration — skipping ${overlayTrackName} alignment`);
    return;
  }

  const overlayGroup = (overlayTrack.medias ?? []).find(
    (m) => m._type === "Group" && overlayIdentRegex.test(m.attributes?.ident ?? "")
  );
  if (!overlayGroup) {
    log("warn", `${overlayTrackName} overlay group not found — skipping`);
    return;
  }

  const oldDuration = overlayGroup.duration ?? 0;
  const oldEnd = (overlayGroup.start ?? 0) + oldDuration;
  const newDuration = sourceEnd - (overlayGroup.start ?? 0);
  if (newDuration <= 0) {
    log("warn", `${overlayTrackName} starts after ${sourceTrackName} ends — skipping`);
    return;
  }

  overlayGroup.duration = newDuration;
  if (typeof overlayGroup.mediaDuration === "number") {
    overlayGroup.mediaDuration = newDuration;
  }

  let stretchedChildren = 0;
  for (const innerTrack of overlayGroup.tracks ?? []) {
    for (const innerMedia of innerTrack.medias ?? []) {
      const innerStart = innerMedia.start ?? 0;
      const innerDuration = newDuration - innerStart;
      if (innerDuration > 0) {
        innerMedia.duration = innerDuration;
        if (typeof innerMedia.mediaDuration === "number") {
          innerMedia.mediaDuration = innerDuration;
        }
        stretchedChildren++;
      }
    }
  }

  log("ok", `${overlayTrackName} aligned to ${sourceTrackName} length`);
  log("detail", `was: ${ticksToTimestamp(overlayGroup.start ?? 0)} → ${ticksToTimestamp(oldEnd)}`);
  log("detail", `now: ${ticksToTimestamp(overlayGroup.start ?? 0)} → ${ticksToTimestamp(sourceEnd)}`);
  log("detail", `inner medias stretched: ${stretchedChildren}`);
}

function alignHostGuestNameOverlays(tracks, trackAttributes) {
  alignNameOverlayToTrackLength(tracks, trackAttributes, "Host", "HostName", /HostNameOverlay/i);
  alignNameOverlayToTrackLength(tracks, trackAttributes, "Guest", "GuestName", /GuestNameOverlay/i);
}

function transitionRefsExist(transition, idSet) {
  for (const key of ["leftMedia", "rightMedia"]) {
    if (key in transition && typeof transition[key] === "number" && !idSet.has(transition[key])) {
      return false;
    }
  }
  return true;
}

function cloneTransitionWithRefs(template, leftMedia, rightMedia) {
  const tr = JSON.parse(JSON.stringify(template));
  if (leftMedia === undefined) {
    delete tr.leftMedia;
  } else {
    tr.leftMedia = leftMedia;
  }
  if (rightMedia === undefined) {
    delete tr.rightMedia;
  } else {
    tr.rightMedia = rightMedia;
  }
  return tr;
}

function applyLogoOutroSequence(tracks, counter, contentDuration) {
  const glowName = /mtr glow transparent/i;

  // Locate the track that contains the logo glow sequence.
  let logoTrack = null;
  let bestCount = 0;
  for (const track of tracks) {
    const count = track.medias.filter(
      (m) => m._type === "IMFile" && glowName.test(m.attributes?.ident ?? "")
    ).length;
    if (count > bestCount) {
      bestCount = count;
      logoTrack = track;
    }
  }

  if (!logoTrack || bestCount < 2) {
    log("warn", "Logo sequence not found (expected at least 2 'mtr glow transparent' clips). Skipping logo automation.");
    return;
  }

  const originalTransitions = logoTrack.transitions ?? [];

  const glowMedias = logoTrack.medias
    .filter((m) => m._type === "IMFile" && glowName.test(m.attributes?.ident ?? ""))
    .sort((a, b) => (a.start ?? 0) - (b.start ?? 0));

  const introTemplate = glowMedias[0];
  const finalTemplate = glowMedias[glowMedias.length - 1];
  const middleTemplate = glowMedias.length >= 3 ? glowMedias[1] : glowMedias[0];

  const finalStart = contentDuration;
  const introClone = cloneWithNewIds(introTemplate, counter);
  introClone.start = introTemplate.start;

  const newGlowMedias = [introClone];
  const templateStart = middleTemplate.start ?? 0;
  const templateDuration = middleTemplate.duration ?? 0;

  if (templateDuration > 0 && finalStart > templateStart) {
    let t = templateStart;
    while (t < finalStart) {
      const remaining = finalStart - t;
      const sliceDuration = Math.min(templateDuration, remaining);
      const middleClone = cloneWithNewIds(middleTemplate, counter);
      middleClone.start = t;
      middleClone.duration = sliceDuration;
      newGlowMedias.push(middleClone);
      t += sliceDuration;
    }
  }

  const finalClone = cloneWithNewIds(finalTemplate, counter);
  finalClone.start = finalStart;
  newGlowMedias.push(finalClone);

  const nonGlowMedias = logoTrack.medias.filter(
    (m) => !(m._type === "IMFile" && glowName.test(m.attributes?.ident ?? ""))
  );

  logoTrack.medias = [...nonGlowMedias, ...newGlowMedias].sort((a, b) => (a.start ?? 0) - (b.start ?? 0));

  // Rebuild transitions so repeated middle clips keep Glow transitions.
  const glowCrossTemplate = originalTransitions.find(
    (t) => t.name === "Glow" && typeof t.leftMedia === "number" && typeof t.rightMedia === "number"
  ) ?? {
    name: "Glow",
    duration: 352800000,
    attributes: {
      bypass: false,
      reverse: false,
      trivial: false,
      useAudioPreRoll: true,
      useVisualPreRoll: true,
    },
  };

  const glowInTemplate = originalTransitions.find(
    (t) => t.name === "Glow" && typeof t.rightMedia === "number" && typeof t.leftMedia !== "number"
  );

  const transmissionTemplate = originalTransitions.find(
    (t) => t.name === "Transmission"
  );

  const rebuiltTransitions = [];

  // Preserve intro glow-in if the project had one.
  if (glowInTemplate) {
    rebuiltTransitions.push(cloneTransitionWithRefs(glowInTemplate, undefined, introClone.id));
  }

  // Add Glow between every adjacent logo segment (intro -> middle... -> final).
  for (let i = 0; i < newGlowMedias.length - 1; i++) {
    rebuiltTransitions.push(
      cloneTransitionWithRefs(glowCrossTemplate, newGlowMedias[i].id, newGlowMedias[i + 1].id)
    );
  }

  // Preserve the final outro transmission behavior.
  if (transmissionTemplate) {
    rebuiltTransitions.push(cloneTransitionWithRefs(transmissionTemplate, finalClone.id, undefined));
  }

  const idSet = new Set(logoTrack.medias.map((m) => m.id));
  logoTrack.transitions = rebuiltTransitions.filter((tr) => transitionRefsExist(tr, idSet));

  const middleCount = Math.max(0, newGlowMedias.length - 2);
  log(
    "ok",
    `Logo track ${logoTrack.trackIndex}: built sequence with ${middleCount} middle repeat(s), final starts at ${ticksToTimestamp(finalStart)}`
  );
}

/**
 * Reposition the outro sound to end exactly when the logo animation ends.
 * Finds the outro media on the outro track and adjusts its start time so that
 * its end matches the logo animation's end time.
 */
function repositionOutroToLogoEnd(tracks) {
  const glowName = /mtr glow transparent/i;

  // Find the logo track
  let logoTrack = null;
  for (const track of tracks) {
    const glowCount = track.medias.filter((m) => m._type === "IMFile" && glowName.test(m.attributes?.ident ?? "")).length;
    if (glowCount > 0) {
      if (!logoTrack || glowCount > (logoTrack._glowCount || 0)) {
        logoTrack = track;
        logoTrack._glowCount = glowCount;
      }
    }
  }

  if (!logoTrack) {
    log("warn", "Logo track not found — cannot reposition outro");
    return;
  }

  // Find logo end time (latest glow segment end)
  const glowMedias = logoTrack.medias.filter((m) => m._type === "IMFile" && glowName.test(m.attributes?.ident ?? ""));
  if (glowMedias.length === 0) {
    log("warn", "No glow media found in logo track — cannot reposition outro");
    return;
  }

  let logoEnd = 0;
  for (const glow of glowMedias) {
    const glowEnd = (glow.start ?? 0) + (glow.duration ?? 0);
    if (glowEnd > logoEnd) logoEnd = glowEnd;
  }

  // Find the outro track and media
  let outroTrack = null;
  let outroMedia = null;

  for (const track of tracks) {
    for (const media of track.medias) {
      if (/outro/i.test(media.attributes?.ident ?? "")) {
        outroTrack = track;
        outroMedia = media;
        break;
      }
    }
    if (outroMedia) break;
  }

  if (!outroMedia) {
    log("warn", "Outro media not found — skipping repositioning");
    return;
  }

  if (!outroTrack) {
    log("warn", "Outro track not found — skipping repositioning");
    return;
  }

  // Calculate new start position: logoEnd - outroDuration
  const newStart = logoEnd - (outroMedia.duration ?? 0);
  const oldStart = outroMedia.start ?? 0;
  const oldEnd = oldStart + (outroMedia.duration ?? 0);
  const newEnd = newStart + (outroMedia.duration ?? 0);

  outroMedia.start = newStart;

  log("ok", `Outro repositioned:`);
  log("detail", `was:     ${ticksToTimestamp(oldStart)} → ${ticksToTimestamp(oldEnd)}`);
  log("detail", `now:     ${ticksToTimestamp(newStart)} → ${ticksToTimestamp(newEnd)}`);
  log("detail", `aligned: outro ends at ${ticksToTimestamp(logoEnd)} (logo animation end)`);
}

/**
 * Reposition the 3rd glitch sound to start immediately after the outro ends.
 * Only the 3rd occurrence (index 2) on the glitch track is moved.
 */
function repositionThirdGlitchAfterOutro(tracks) {
  // Find outro end time
  let outroEnd = 0;
  for (const track of tracks) {
    for (const media of track.medias) {
      if (/outro/i.test(media.attributes?.ident ?? "")) {
        outroEnd = (media.start ?? 0) + (media.duration ?? 0);
        break;
      }
    }
    if (outroEnd > 0) break;
  }

  if (outroEnd === 0) {
    log("warn", "Outro not found — cannot reposition 3rd glitch");
    return;
  }

  // Find the glitch track and collect all glitch sounds
  for (const track of tracks) {
    const glitchMedias = track.medias.filter((m) => /glitch/i.test(m.attributes?.ident ?? ""));
    if (glitchMedias.length < 3) continue;

    const third = glitchMedias[2];
    const oldStart = third.start ?? 0;
    third.start = outroEnd;

    log("ok", `3rd glitch sound repositioned:`);
    log("detail", `was:  ${ticksToTimestamp(oldStart)} → ${ticksToTimestamp(oldStart + (third.duration ?? 0))}`);
    log("detail", `now:  ${ticksToTimestamp(third.start)} → ${ticksToTimestamp(third.start + (third.duration ?? 0))}`);
    log("detail", `starts right after outro end (${ticksToTimestamp(outroEnd)})`);
    return;
  }

  log("warn", "Could not find a track with 3 or more glitch sounds — skipping");
}

/**
 * Stretch the stars background so it ends at the same time as the 3rd glitch sound.
 * The 3rd glitch end is: outroEnd + glitch duration.
 */
function stretchStarsBGToGlitchEnd(tracks) {
  // Find the 3rd glitch start time (= outro end)
  let glitchStart = 0;
  for (const track of tracks) {
    const glitchMedias = track.medias.filter((m) => /glitch/i.test(m.attributes?.ident ?? ""));
    if (glitchMedias.length >= 3) {
      glitchStart = glitchMedias[2].start ?? 0;
      break;
    }
  }

  if (glitchStart === 0) {
    log("warn", "3rd glitch not found — cannot stretch stars BG");
    return;
  }

  // Find the stars media and its outgoing transition duration
  for (const track of tracks) {
    for (const media of track.medias) {
      if (/\bstars\b/i.test(media.attributes?.ident ?? "")) {
        // Find the outgoing transition (leftMedia === media.id)
        const outgoingTransition = (track.transitions ?? []).find(
          (tr) => tr.leftMedia === media.id
        );
        const transitionDuration = outgoingTransition?.duration ?? 0;

        const oldDuration = media.duration ?? 0;
        const oldEnd = (media.start ?? 0) + oldDuration;
        // Stars ends at glitch_start + transition_duration so the transition starts exactly at glitch_start
        const newEnd = glitchStart + transitionDuration;
        const newDuration = newEnd - (media.start ?? 0);
        media.duration = newDuration;

        log("ok", `Stars BG stretched so outgoing transition starts at 3rd glitch start:`);
        log("detail", `was:  ${ticksToTimestamp(media.start ?? 0)} → ${ticksToTimestamp(oldEnd)}`);
        log("detail", `now:  ${ticksToTimestamp(media.start ?? 0)} → ${ticksToTimestamp(newEnd)}`);
        log("detail", `transition start: ${ticksToTimestamp(glitchStart)} (duration: ${ticksToTimestamp(transitionDuration)})`);
        return;
      }
    }
  }

  log("warn", "Stars media not found — skipping");
}

/**
 * Extend the EpisodeNameOverlay group so it ends at the start of the final
 * (animated) logo glow clip — i.e. exactly when the logo animation begins.
 */
function extendEpisodeNameOverlay(tracks) {
  const glowName = /mtr glow transparent/i;

  // Find the logo track and its final glow start
  let finalGlowStart = 0;
  for (const track of tracks) {
    const glows = track.medias.filter((m) => m._type === "IMFile" && glowName.test(m.attributes?.ident ?? ""));
    if (glows.length > 0) {
      finalGlowStart = glows[glows.length - 1].start ?? 0;
      break;
    }
  }

  if (finalGlowStart === 0) {
    log("warn", "Logo glow not found — cannot extend EpisodeNameOverlay");
    return;
  }

  // Find EpisodeNameOverlay
  for (const track of tracks) {
    for (const media of track.medias) {
      if (/EpisodeNameOverlay/i.test(media.attributes?.ident ?? "")) {
        const oldDuration = media.duration ?? 0;
        const oldEnd = (media.start ?? 0) + oldDuration;
        const newDuration = finalGlowStart - (media.start ?? 0);

        if (newDuration <= 0) {
          log("warn", `EpisodeNameOverlay already ends after logo start — skipping`);
          return;
        }

        media.duration = newDuration;

        // Pass 1: stretch all inner callouts to match the final group length.
        let updatedCallouts = 0;
        for (const innerTrack of media.tracks ?? []) {
          for (const innerMedia of innerTrack.medias ?? []) {
            if (innerMedia._type !== "Callout") continue;

            const innerStart = innerMedia.start ?? 0;
            const innerDuration = media.duration - innerStart;
            if (innerDuration > 0) {
              innerMedia.duration = innerDuration;
              if (typeof innerMedia.mediaDuration === "number") {
                innerMedia.mediaDuration = innerDuration;
              }
            }

            updatedCallouts++;
          }
        }

        log("ok", `EpisodeNameOverlay extended to logo animation start:`);
        log("detail", `was:  ${ticksToTimestamp(media.start ?? 0)} → ${ticksToTimestamp(oldEnd)}`);
        log("detail", `now:  ${ticksToTimestamp(media.start ?? 0)} → ${ticksToTimestamp(finalGlowStart)}`);
        log("detail", `inner callouts stretched: ${updatedCallouts} (duration)`);
        return;
      }
    }
  }

  log("warn", "EpisodeNameOverlay not found — skipping");
}

/**
 * Pass 2 for EpisodeNameOverlay:
 * apply scale behavior after durations are finalized so behavior timing uses
 * the final callout lengths.
 */
function applyScaleBehaviorToEpisodeNameOverlay(tracks) {
  function buildScaleBehaviorEffect(duration) {
    return {
      _type: "GenericBehaviorEffect",
      effectName: "scale",
      bypassed: false,
      start: 0,
      duration,
      in: {
        attributes: {
          name: "grow",
          type: 1,
          characterOrder: 6,
          offsetBetweenCharacters: 23520000,
          suggestedDurationPerCharacter: 423360000,
          overlapProportion: "3/4",
          movement: 30,
          springDamping: 5,
          springStiffness: 50,
          bounceBounciness: 0.45,
        },
      },
      center: {
        attributes: {
          name: "none",
          type: 1,
          characterOrder: 6,
          offsetBetweenCharacters: 0,
          secondsPerLoop: 1,
          numberOfLoops: -1,
          delayBetweenLoops: 0,
        },
      },
      out: {
        attributes: {
          name: "shrink",
          type: 1,
          characterOrder: 6,
          offsetBetweenCharacters: 23520000,
          suggestedDurationPerCharacter: 341040000,
          overlapProportion: "1/10",
          movement: 11,
          springDamping: 5,
          springStiffness: 50,
          bounceBounciness: 0.45,
        },
      },
    };
  }

  for (const track of tracks) {
    for (const media of track.medias) {
      if (!/EpisodeNameOverlay/i.test(media.attributes?.ident ?? "")) continue;

      let updatedCallouts = 0;
      for (const innerTrack of media.tracks ?? []) {
        for (const innerMedia of innerTrack.medias ?? []) {
          if (innerMedia._type !== "Callout") continue;

          if (!innerMedia.metadata || typeof innerMedia.metadata !== "object") {
            innerMedia.metadata = {};
          }
          innerMedia.metadata.effectApplied = "scale";
          if (!("default-scale" in innerMedia.metadata)) {
            innerMedia.metadata["default-scale"] = "1";
          }

          const finalDuration = innerMedia.duration ?? 0;

          // Replace any existing behavior effects and apply the exact scale template.
          const keptEffects = (innerMedia.effects ?? []).filter(
            (e) => e?._type !== "GenericBehaviorEffect"
          );
          keptEffects.push(buildScaleBehaviorEffect(finalDuration));
          innerMedia.effects = keptEffects;

          if (!innerMedia.parameters || typeof innerMedia.parameters !== "object") {
            innerMedia.parameters = {};
          }

          updatedCallouts++;
        }
      }

      log("ok", `EpisodeNameOverlay scale behavior applied:`);
      log("detail", `callouts updated: ${updatedCallouts}`);
      return;
    }
  }

  log("warn", "EpisodeNameOverlay not found for scale behavior — skipping");
}

function listAllGroups(tracks, trackAttributes) {
  log("step", "All top-level Groups found in this project:");
  let count = 0;
  for (const track of tracks) {
    const trackLabel = trackAttributes?.[track.trackIndex]?.ident || "unlabeled";
    for (const media of track.medias) {
      if (media._type === "Group") {
        const ident = media.attributes?.ident ?? "(no ident)";
        const dur = ticksToTimestamp(media.duration);
        const start = ticksToTimestamp(media.start);
        log(
          "info",
          `track ${String(track.trackIndex).padEnd(3)} (${trackLabel}) | ` +
          `id ${String(media.id).padEnd(4)} | ` +
          `starts ${start} | duration ${dur} | ` +
          `ident: "${ident}"`
        );
        count++;
      }
    }
  }
  if (count === 0) {
    log("warn", "No Groups found at the top level.");
    log("warn", "Make sure the path points to project.tscproj inside the .cmproj bundle.");
  }
}

// ---------------------------------------------------------------------------
// CLONING
// ---------------------------------------------------------------------------

function cloneWithNewIds(media, counter) {
  const clone = JSON.parse(JSON.stringify(media));
  const idMap = new Map();

  function reassign(m) {
    const oldId = m.id;
    const newId = counter.next++;
    m.id = newId;
    idMap.set(oldId, newId);
    if (m._type === "Group" && m.tracks) {
      for (const track of m.tracks) {
        for (const child of track.medias) {
          reassign(child);
        }
      }
    }
  }

  function remapKnownRefs(node, parentKey = "") {
    if (Array.isArray(node)) {
      if (parentKey === "objects") {
        for (let i = 0; i < node.length; i++) {
          const item = node[i];
          if (typeof item === "number" && idMap.has(item)) {
            node[i] = idMap.get(item);
          } else {
            remapKnownRefs(item, "objects[]");
          }
        }
      } else {
        for (const item of node) {
          remapKnownRefs(item);
        }
      }
      return;
    }

    if (!node || typeof node !== "object") return;

    for (const [key, value] of Object.entries(node)) {
      if (key === "id") continue;

      if ((key === "leftMedia" || key === "rightMedia" || key === "media") && typeof value === "number" && idMap.has(value)) {
        node[key] = idMap.get(value);
        continue;
      }

      remapKnownRefs(value, key);
    }
  }

  reassign(clone);
  remapKnownRefs(clone);
  return clone;
}

// ---------------------------------------------------------------------------
// INJECTION
// ---------------------------------------------------------------------------

function injectOverlays(projectPath, rules) {

  // Step 1: Load
  log("step", `Loading: ${projectPath}`);
  let raw;
  try {
    raw = fs.readFileSync(projectPath, "utf-8");
  } catch (e) {
    log("error", `Cannot read file: ${e.message}`);
    process.exit(1);
  }
  log("ok", `File read (${(raw.length / 1024).toFixed(1)} KB)`);

  let project;
  try {
    project = JSON.parse(raw);
  } catch (e) {
    log("error", `JSON parse failed: ${e.message}`);
    process.exit(1);
  }
  log("ok", "JSON parsed successfully");

  // Step 2: Read structure
  log("step", "Reading timeline structure");

  let tracks;
  try {
    tracks = getSceneTracks(project);
  } catch (e) {
    log("error", `Could not read timeline tracks: ${e.message}`);
    process.exit(1);
  }

  const trackAttributes = getTrackAttributes(project);

  log("info", `Found ${tracks.length} top-level track(s):`);
  for (const t of tracks) {
    const label = trackAttributes[t.trackIndex]?.ident || "(unlabeled)";
    log(
      "sub",
      `trackIndex ${String(t.trackIndex).padEnd(3)} | "${label}" | ${t.medias.length} media(s)`
    );
  }

  // Step 2b: Stretch HostNameOverlay/GuestNameOverlay to Host/Guest durations.
  log("step", "Aligning host/guest name overlays");
  alignHostGuestNameOverlays(tracks, trackAttributes);

  const videoDuration = getVideoDuration(tracks);
  const videoDurationSec = (videoDuration / TICKS_PER_SECOND).toFixed(1);
  log("ok", `Video duration: ${ticksToTimestamp(videoDuration)} (${videoDurationSec}s)`);
  const contentDuration = getContentDuration(tracks);
  const contentDurationSec = (contentDuration / TICKS_PER_SECOND).toFixed(1);
  log(
    "ok",
    `Content duration: ${ticksToTimestamp(contentDuration)} (${contentDurationSec}s)` +
    (contentDuration === videoDuration ? " (same as video)" : " (from host/guest tracks)")
  );

  const maxId = findMaxId(tracks);
  const counter = { next: maxId + 1 };
  log("info", `Highest existing media ID: ${maxId} — new IDs start from ${counter.next}`);

  // Step 3: Inventory groups
  listAllGroups(tracks, trackAttributes);

  // Step 4: Process rules
  log("step", `Processing ${rules.length} rule(s)`);
  let totalInjected = 0;

  // Build mutex groups for cross-track deconfliction (e.g. CTA overlays).
  const mutexTracksByGroup = new Map();
  for (const rule of rules) {
    if (!rule.mutexGroup) continue;
    const track = tracks.find((t) => t.trackIndex === rule.trackIndex);
    if (!track) continue;
    if (!mutexTracksByGroup.has(rule.mutexGroup)) {
      mutexTracksByGroup.set(rule.mutexGroup, []);
    }
    const arr = mutexTracksByGroup.get(rule.mutexGroup);
    if (!arr.includes(track)) arr.push(track);
  }

  const deconflictStep = secondsToTicks(5);

  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    console.log(`\n  -- Rule ${i + 1}/${rules.length}: "${rule.ident}" -> track ${rule.trackIndex} --`);

    log("info", `Searching for Group with ident "${rule.ident}"...`);
    const found = findOverlayByIdent(tracks, rule.ident, rule.trackIndex);
    if (!found) {
      log("warn", "Not found — skipping. Check the ident strings printed above.");
      continue;
    }
    const { media: source, track: sourceTrack } = found;
    log(
      "ok",
      `Found — id: ${source.id}, on track: ${sourceTrack.trackIndex}, duration: ${ticksToTimestamp(source.duration)}`
    );

    const targetTrack = tracks.find((t) => t.trackIndex === rule.trackIndex);
    if (!targetTrack) {
      log("warn", `Track index ${rule.trackIndex} not found. Available: ${tracks.map((t) => t.trackIndex).join(", ")}`);
      continue;
    }
    log("ok", `Target track ${rule.trackIndex} found (${targetTrack.medias.length} existing media(s))`);

    // Keep one template instance and remove previously injected instances for this ident
    // so each run rebuilds a clean, deterministic schedule.
    const beforeCount = targetTrack.medias.length;
    targetTrack.medias = targetTrack.medias.filter((m) => {
      if (m._type !== "Group") return true;
      if (m.attributes?.ident !== rule.ident) return true;
      return m.id === source.id;
    });


    const removed = beforeCount - targetTrack.medias.length;
    if (removed > 0) {
      log("info", `Removed ${removed} previous injected instance(s) for "${rule.ident}"`);
    }

    let timesTicks;
    let placementCutoff = contentDuration - source.duration;

    if (Array.isArray(rule.timesSeconds)) {
      timesTicks = rule.timesSeconds.map(secondsToTicks);
      log("info", `Mode: explicit — ${timesTicks.length} timestamp(s)`);
    } else {
      const interval = secondsToTicks(rule.intervalSeconds);
      const offset = secondsToTicks(rule.offsetSeconds ?? 0);
      const skipEnd = secondsToTicks(rule.skipLastSeconds ?? 0);
      const cutoff = contentDuration - skipEnd - source.duration;
      placementCutoff = cutoff;

      log("info", "Mode: interval");
      log("detail", `every:      ${ticksToTimestamp(interval)} (${rule.intervalSeconds}s)`);
      log("detail", `first at:   ${ticksToTimestamp(offset)} (offset ${rule.offsetSeconds ?? 0}s)`);
      log("detail", `last valid: ${ticksToTimestamp(cutoff)} (skip last ${rule.skipLastSeconds ?? 0}s + overlay duration)`);

      timesTicks = [];
      let t = offset;
      while (t <= cutoff) {
        timesTicks.push(t);
        t += interval;
      }

      if (timesTicks.length === 0) {
        log("warn", "No placements generated — window may be too small.");
        log(
          "warn",
          `content: ${contentDurationSec}s | offset: ${rule.offsetSeconds ?? 0}s | ` +
          `skip: ${rule.skipLastSeconds ?? 0}s | overlay: ${(source.duration / TICKS_PER_SECOND).toFixed(1)}s`
        );
        continue;
      }
    }

    log("ok", `Attempting ${timesTicks.length} placement(s):`);
    let placed = 0;
    const mutexTracks = rule.mutexGroup ? (mutexTracksByGroup.get(rule.mutexGroup) ?? []) : [];
    const mutexTracksExcludingTarget = mutexTracks.filter((t) => t !== targetTrack);

    for (const desiredStart of timesTicks) {
      let t = desiredStart;
      let targetCollision = findCollisionInTrack(targetTrack, t, source.duration);
      let groupCollision = findCollisionInTrackSet(mutexTracksExcludingTarget, t, source.duration);

      // For mutex-group overlays, push start forward until there is no overlap
      // with same-track or grouped-track media.
      if (rule.mutexGroup) {
        while ((targetCollision || groupCollision) && t <= placementCutoff) {
          t += deconflictStep;
          targetCollision = findCollisionInTrack(targetTrack, t, source.duration);
          groupCollision = findCollisionInTrackSet(mutexTracksExcludingTarget, t, source.duration);
        }
      }

      if (t > placementCutoff || targetCollision || groupCollision) {
        if (groupCollision) {
          log(
            "warn",
            `Skipping ${ticksToTimestamp(desiredStart)} due to overlap with group track ${groupCollision.track.trackIndex} media id ${groupCollision.media.id}`
          );
        } else if (targetCollision) {
          log(
            "warn",
            `Skipping ${ticksToTimestamp(desiredStart)} due to overlap with existing media id ${targetCollision.id} ` +
            `(${ticksToTimestamp(targetCollision.start)} to ${ticksToTimestamp((targetCollision.start ?? 0) + (targetCollision.duration ?? 0))})`
          );
        } else {
          log("warn", `Skipping ${ticksToTimestamp(desiredStart)} — no deconflicted slot before cutoff`);
        }
        continue;
      }

      if (t !== desiredStart) {
        log("detail", `shifted ${ticksToTimestamp(desiredStart)} -> ${ticksToTimestamp(t)} to avoid cross-track overlap`);
      }

      log("sub", `${ticksToTimestamp(t)}  (tick ${t})  — IDs from ${counter.next}`);
      const clone = cloneWithNewIds(source, counter);
      clone.start = t;
      targetTrack.medias.push(clone);
      placed++;
    }

    if (placed === 0) {
      log("warn", "No non-overlapping placements were possible for this rule.");
    } else {
      log("ok", `Placed ${placed} instance(s) for this rule`);
    }
    log("ok", `Track ${rule.trackIndex} now has ${targetTrack.medias.length} media(s)`);
    totalInjected += placed;
  }

  // Final guard: ensure mutex-group overlays never play simultaneously.
  enforceMutexGroupNoOverlap(tracks, rules);

  // Step 4b: Build logo outro sequence.
  // Repeat the middle glow clip continuously up to content end,
  // then place the final animated glow clip at content end.
  log("step", "Applying logo outro automation");
  applyLogoOutroSequence(tracks, counter, contentDuration);

  // Step 4c: Reposition outro sound to end with logo animation.
  log("step", "Repositioning outro sound");
  repositionOutroToLogoEnd(tracks);

  // Step 4d: Reposition 3rd glitch sound to start after outro ends.
  log("step", "Repositioning 3rd glitch sound");
  repositionThirdGlitchAfterOutro(tracks);

  // Step 4e: Stretch stars BG to end at the same time as the 3rd glitch sound.
  log("step", "Stretching stars background");
  stretchStarsBGToGlitchEnd(tracks);

  // Step 4f: Extend EpisodeNameOverlay to reach the start of the logo animation.
  log("step", "Extending EpisodeNameOverlay");
  extendEpisodeNameOverlay(tracks);

  // Step 4g: Apply scale behavior to EpisodeName callouts after final lengths are known.
  log("step", "Applying EpisodeName scale behavior");
  applyScaleBehaviorToEpisodeNameOverlay(tracks);

  // Step 5: Write output
  //
  // Input/output: episode.cmproj/project.tscproj
  //
  // We update the tscproj in place and keep a .bak backup beside it.
  //
  log("step", "Writing output");

  const tscprojAbs = path.resolve(projectPath);
  const bundlePath = path.dirname(tscprojAbs);             // .../episode.cmproj
  const outTscproj = tscprojAbs;
  const backupPath = `${outTscproj}.bak`;

  log("info", `Project bundle : ${bundlePath}`);
  log("info", `Target tscproj : ${outTscproj}`);

  if (fs.existsSync(outTscproj)) {
    fs.copyFileSync(outTscproj, backupPath);
    log("info", `Backed up previous tscproj -> ${path.basename(backupPath)}`);
  }

  // Write the modified project.tscproj into the copy
  // Camtasia uses int64 sentinel values (e.g. 9223372036854775807).
  // JS Number cannot represent them exactly and rounds during parse/stringify.
  // Repair known rounded artifacts before saving.
  const outJson = JSON.stringify(project, null, 2)
    .replaceAll(JS_ROUNDED_INT64_MAX, CAMTASIA_INT64_MAX);

  fs.writeFileSync(outTscproj, outJson, "utf-8");
  log("ok", `Written: ${outTscproj} (${(outJson.length / 1024).toFixed(1)} KB)`);

  // Summary
  console.log("\n=============================================");
  console.log(`  Rules processed : ${rules.length}`);
  console.log(`  Total injected  : ${totalInjected} overlay instance(s)`);
  console.log(`  Updated project : ${outTscproj}`);
  console.log(`  Backup file     : ${backupPath}`);
  console.log("=============================================");
  console.log(`\n  Open in Camtasia:\n  open "${bundlePath}"\n`);
}

// ---------------------------------------------------------------------------
// RULES — edit these to match your episode
// ---------------------------------------------------------------------------
//
// Each rule needs:
//   ident       — exact "ident" string from the Groups inventory printed at runtime
//   trackIndex  — which top-level track to inject into (also printed at runtime)
//
// Interval mode:
//   intervalSeconds   — repeat every N seconds
//   offsetSeconds     — (optional) first placement time in seconds (default: 0)
//   skipLastSeconds   — (optional) don't place within N seconds of end (default: 0)
//
// Explicit mode:
//   timesSeconds      — array of exact timestamps, e.g. [60, 300, 600]
//
// ---------------------------------------------------------------------------

const rules = [
  // Tagline bar — every 5 min. The template on track 8 is the first appearance;
  // repeats start at the first interval mark and skip the last 1 min of content.
  {
    ident: "TaglineOverlay",
    trackIndex: 8,
    intervalSeconds: 300,
    offsetSeconds: 300,
    skipLastSeconds: 60,
    mutexGroup: "cta-overlays",
  },

  // Subscribe CTA — every 5 min. First repeat at 5:00, skip last 2 min.
  {
    ident: "Subscribe",
    trackIndex: 9,
    intervalSeconds: 300,
    offsetSeconds: 300,
    skipLastSeconds: 120,
    mutexGroup: "cta-overlays",
  },

  // Website lower third — every 8 min. First repeat at 8:00, skip last 2 min.
  {
    ident: "WebsiteOverlay",
    trackIndex: 10,
    intervalSeconds: 480,
    offsetSeconds: 480,
    skipLastSeconds: 120,
    mutexGroup: "cta-overlays",
  },

  // Socials bar — every 7 min. First repeat at 7:00, skip last 1 min.
  {
    ident: "SocialsOverlay",
    trackIndex: 11,
    intervalSeconds: 420,
    offsetSeconds: 420,
    skipLastSeconds: 60,
    mutexGroup: "cta-overlays",
  },
];

// ---------------------------------------------------------------------------
// ENTRY POINT
// ---------------------------------------------------------------------------

const projectPath = process.argv[2];

console.log("\n=============================================");
console.log("  Camtasia Overlay Injector");
console.log("=============================================");

if (!projectPath) {
  console.log("  Usage: node episode-injector.js <path/to/project.tscproj>");
  console.log("  Example: node episode-injector.js episode.cmproj/project.tscproj\n");
  process.exit(1);
}

if (!fs.existsSync(projectPath)) {
  log("error", `File not found: ${projectPath}`);
  process.exit(1);
}

console.log(`  Input : ${projectPath}`);
console.log(`  Rules : ${rules.length}`);
console.log("=============================================");

injectOverlays(projectPath, rules);