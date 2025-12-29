/**
 * Suffix Configuration Service
 * @fileoverview Canonical suffix map and helper functions for event classification
 * Version 7.9.7+
 *
 * This module defines the single source of truth for all event suffixes A-Z,
 * including Commander brackets, Limited formats, and mission tags.
 */

// ============================================================================
// COSMIC SUFFIX MAP v7.9.7+
// ============================================================================

/**
 * SUFFIX_MAP - Single source of truth for all event suffixes.
 *
 * Keys:
 *   code              - Single-letter suffix A–Z
 *   name              - Human-readable name
 *   game              - Primary game/line (or "MULTI" / "GENERIC")
 *   formatType        - "CONSTRUCTED" | "LIMITED" | "PROGRAM" | "HOBBY" | "INTERNAL"
 *   commanderBracket  - null | 1 | 2 | 3 | 4 | 5  (top bracket if range)
 *   commanderRange    - null | [minBracket, maxBracket]
 *   requiresKitPrompt - true if D/R/S/P RL95 needs kit-cost prompt
 *   missionTags       - Array of tags used by MissionLog / KPI engine
 */
const SUFFIX_MAP = {
  "A": {
    code: "A",
    name: "Academy / Learn to Play",
    game: "MULTI",
    formatType: "PROGRAM",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["ACADEMY", "ONBOARDING", "NEW_PLAYER"]
  },
  "B": {
    code: "B",
    name: "Casual Commander (Brk 1–2)",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: 2,
    commanderRange: [1, 2],
    requiresKitPrompt: false,
    missionTags: ["COMMANDER", "BRK_1_2", "CASUAL"]
  },
  "C": {
    code: "C",
    name: "Transitional Commander (Brk 3–4)",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: 4,
    commanderRange: [3, 4],
    requiresKitPrompt: false,
    missionTags: ["COMMANDER", "BRK_3_4", "TRANSITION"]
  },
  "D": {
    code: "D",
    name: "Booster Draft",
    game: "MTG",
    formatType: "LIMITED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: true,
    missionTags: ["MTG", "DRAFT", "LIMITED"]
  },
  "E": {
    code: "E",
    name: "External / Outreach",
    game: "MULTI",
    formatType: "PROGRAM",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["OUTREACH", "OFFSITE", "MARKETING"]
  },
  "F": {
    code: "F",
    name: "Free Play Event",
    game: "MULTI",
    formatType: "PROGRAM",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["FREE_PLAY", "COMMUNITY"]
  },
  "G": {
    code: "G",
    name: "Gundam / Gunpla",
    game: "GUNPLA",
    formatType: "HOBBY",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["GUNPLA", "HOBBY"]
  },
  "H": {
    code: "H",
    name: "Historic / Legacy MTG",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["MTG", "HISTORIC_LEGACY"]
  },
  "I": {
    code: "I",
    name: "Yu-Gi-Oh TCG",
    game: "YUGIOH",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["YUGIOH"]
  },
  "J": {
    code: "J",
    name: "Junior / Youth Events",
    game: "MULTI",
    formatType: "PROGRAM",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["JUNIOR", "YOUTH"]
  },
  "K": {
    code: "K",
    name: "Kill Team",
    game: "KILL_TEAM",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["MINIATURES", "KILL_TEAM"]
  },
  "L": {
    code: "L",
    name: "Commander League",
    game: "MTG",
    formatType: "PROGRAM",
    commanderBracket: null, // League can host B/C/T internally
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["COMMANDER", "LEAGUE"]
  },
  "M": {
    code: "M",
    name: "Modern Constructed",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["MTG", "MODERN"]
  },
  "N": {
    code: "N",
    name: "Pokémon TCG",
    game: "POKEMON",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["POKEMON"]
  },
  "O": {
    code: "O",
    name: "One Piece TCG",
    game: "ONE_PIECE",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["ONE_PIECE"]
  },
  "P": {
    code: "P",
    name: "Proxy / Cube Draft",
    game: "MTG",
    formatType: "LIMITED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: true, // limited packs / kit cost behavior
    missionTags: ["MTG", "CUBE", "PROXY", "LIMITED"]
  },
  "Q": {
    code: "Q",
    name: "Precon Event",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["MTG", "PRECON"]
  },
  "R": {
    code: "R",
    name: "Prerelease Sealed",
    game: "MTG",
    formatType: "LIMITED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: true,
    missionTags: ["MTG", "PRERELEASE", "LIMITED"]
  },
  "S": {
    code: "S",
    name: "Sealed",
    game: "MTG",
    formatType: "LIMITED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: true,
    missionTags: ["MTG", "SEALED", "LIMITED"]
  },
  "T": {
    code: "T",
    name: "Two-Headed Giant Commander",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: 5,
    commanderRange: [5, 5],
    requiresKitPrompt: false,
    missionTags: ["COMMANDER", "BRK_1_3", "Casual"]
  },
  "U": {
    code: "U",
    name: "cEDH / High-Power Commander (Brk 5)",
    game: "MTG",
    formatType: "CONSTRUCTED",
    commanderBracket: 5,
    commanderRange: [5, 5],
    requiresKitPrompt: false,
    missionTags: ["COMMANDER", "BRK_5", "CEDH"]
  },
  "V": {
    code: "V",
    name: "Riftbound",
    game: "RIFTBOUND",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["RIFTBOUND"]
  },
  "W": {
    code: "W",
    name: "Workshop / Hobby Night",
    game: "MULTI",
    formatType: "HOBBY",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["HOBBY", "WORKSHOP"]
  },
  "X": {
    code: "X",
    name: "Multi-Event Day",
    game: "MULTI",
    formatType: "PROGRAM",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["MULTI_EVENT", "FESTIVAL"]
  },
  "Y": {
    code: "Y",
    name: "Lorcana TCG",
    game: "LORCANA",
    formatType: "CONSTRUCTED",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["LORCANA"]
  },
  "Z": {
    code: "Z",
    name: "Staff / Internal Use",
    game: "INTERNAL",
    formatType: "INTERNAL",
    commanderBracket: null,
    commanderRange: null,
    requiresKitPrompt: false,
    missionTags: ["STAFF", "INTERNAL"]
  }
};

// ============================================================================
// SUFFIX HELPERS
// ============================================================================

/**
 * Get suffix metadata by code
 * @param {string} code - Single-letter suffix code (A-Z)
 * @return {Object|null} Suffix metadata or null if not found
 */
function getSuffixMeta_(code) {
  if (!code) return null;
  return SUFFIX_MAP[code] || null;
}

/**
 * Parse suffix from an event sheet name.
 * Valid patterns:
 *   MM-DD-YYYY        -> null (no suffix)
 *   MM-DDX-YYYY       -> X (where X is A–Z suffix)
 *
 * Examples:
 *   "11-23-2025"   -> null
 *   "11-23B-2025"  -> "B"
 *   "05-01T-2026"  -> "T"
 *
 * @param {string} eventId - Event sheet name
 * @return {string|null} Suffix letter or null
 */
function getSuffixFromEventId_(eventId) {
  if (!eventId) return null;
  const parts = eventId.split("-");
  if (parts.length !== 3) return null;

  // parts[1] may be "23" or "23B"
  const dayPart = parts[1];
  const maybeSuffix = dayPart.replace(/[0-9]/g, "");
  if (!maybeSuffix) return null;
  if (maybeSuffix.length !== 1) return null;

  // Validate it's uppercase A-Z
  if (!/^[A-Z]$/.test(maybeSuffix)) return null;

  return maybeSuffix;
}

/**
 * Get all valid suffix codes
 * @return {Array<string>} Array of suffix codes A-Z
 */
function getAllSuffixCodes_() {
  return Object.keys(SUFFIX_MAP);
}

/**
 * Get suffix display name for UI dropdowns
 * @param {string} code - Suffix code
 * @return {string} Display string (e.g., "B – Casual Commander (Brk 1–2)")
 */
function getSuffixDisplayName_(code) {
  const meta = getSuffixMeta_(code);
  if (!meta) return code;
  return `${code} – ${meta.name}`;
}

/**
 * Get all suffix options for UI dropdowns
 * @return {Array<Object>} Array of {code, display}
 */
function getSuffixOptions_() {
  return getAllSuffixCodes_().map(code => ({
    code: code,
    display: getSuffixDisplayName_(code)
  }));
}

/**
 * Check if suffix requires kit cost prompt (Limited formats)
 * @param {string} code - Suffix code
 * @return {boolean} True if D/P/R/S
 */
function suffixRequiresKitPrompt_(code) {
  const meta = getSuffixMeta_(code);
  return meta ? !!meta.requiresKitPrompt : false;
}

/**
 * Get Commander bracket range for suffix (if applicable)
 * @param {string} code - Suffix code
 * @return {Array<number>|null} [min, max] or null
 */
function getCommanderBracketRange_(code) {
  const meta = getSuffixMeta_(code);
  return meta ? meta.commanderRange : null;
}

/**
 * Check if suffix is a Commander format
 * @param {string} code - Suffix code
 * @return {boolean} True if B/C/T/L
 */
function isCommanderSuffix_(code) {
  return ['B', 'C', 'T', 'L'].includes(code);
}

/**
 * Check if suffix is Limited format
 * @param {string} code - Suffix code
 * @return {boolean} True if D/P/R/S
 */
function isLimitedSuffix_(code) {
  return ['D', 'P', 'R', 'S'].includes(code);
}