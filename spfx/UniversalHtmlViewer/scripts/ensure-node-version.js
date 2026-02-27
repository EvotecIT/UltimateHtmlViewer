#!/usr/bin/env node

const requiredMin = { major: 22, minor: 14, patch: 0 };
const requiredMaxMajorExclusive = 23;

const versionRaw = process.versions && process.versions.node ? process.versions.node : '';
const parsed = parseVersion(versionRaw);

if (!parsed) {
  fail(
    `Unable to parse Node.js version "${versionRaw || '(empty)'}". ` +
      'Use Node 22.14.0 (or newer 22.x) for this project.',
  );
}

const isSupportedMajor =
  parsed.major === requiredMin.major && parsed.major < requiredMaxMajorExclusive;
const meetsMin =
  parsed.minor > requiredMin.minor ||
  (parsed.minor === requiredMin.minor && parsed.patch >= requiredMin.patch);

if (!isSupportedMajor || !meetsMin) {
  fail(
    `Unsupported Node.js version v${parsed.major}.${parsed.minor}.${parsed.patch}. ` +
      'Required: >=22.14.0 <23.0.0.\n' +
      'Run commands with:\n' +
      '  npx -y -p node@22.14.0 -c "<command>"',
  );
}

process.exit(0);

function parseVersion(value) {
  const match = String(value || '').trim().match(/^(\d+)\.(\d+)\.(\d+)/);
  if (!match) {
    return undefined;
  }

  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
  };
}

function fail(message) {
  // eslint-disable-next-line no-console
  console.error(`[UHV Node Guard] ${message}`);
  process.exit(1);
}
