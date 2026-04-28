'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const projectRoot = path.resolve(__dirname, '..');
const packageConfig = require(path.join(projectRoot, 'config', 'package-solution.json'));
const iconPath = packageConfig.solution && packageConfig.solution.iconPath;

if (!iconPath) {
  process.exit(0);
}

const sharepointRoot = path.join(projectRoot, 'sharepoint');
const iconSourcePath = path.join(sharepointRoot, iconPath);
const packagePath = path.join(sharepointRoot, packageConfig.paths.zippedPackage);
const debugRoot = path.join(sharepointRoot, 'solution', 'debug');
const relsPath = '_rels/ClientSideAssets.xml.rels';
const clientAssetPath = `ClientSideAssets/${path.basename(iconPath)}`;

function ensureIconExists() {
  if (!fs.existsSync(iconSourcePath)) {
    throw new Error(`Catalog icon not found: ${iconSourcePath}`);
  }
}

function addClientAssetRelationship(relsXml) {
  if (relsXml.includes(`Target="/${clientAssetPath}"`)) {
    return relsXml;
  }

  const relationshipIds = Array.from(relsXml.matchAll(/\bId="r(\d+)"/g))
    .map((match) => Number(match[1]))
    .filter((id) => Number.isInteger(id));
  const nextId = relationshipIds.length ? Math.max(...relationshipIds) + 1 : 1;
  const relationship =
    `<Relationship Type="http://schemas.microsoft.com/sharepoint/2012/app/relationships/clientsideasset" ` +
    `Target="/${clientAssetPath}" Id="r${nextId}"></Relationship>`;

  return relsXml.replace('</Relationships>', `${relationship}</Relationships>`);
}

function patchDebugPackage() {
  const debugRelsPath = path.join(debugRoot, relsPath);
  if (!fs.existsSync(debugRelsPath)) {
    return;
  }

  const debugClientAssetPath = path.join(debugRoot, clientAssetPath);
  fs.mkdirSync(path.dirname(debugClientAssetPath), { recursive: true });
  fs.copyFileSync(iconSourcePath, debugClientAssetPath);

  const relsXml = fs.readFileSync(debugRelsPath, 'utf8');
  fs.writeFileSync(debugRelsPath, addClientAssetRelationship(relsXml));
}

async function patchSppkg() {
  if (!fs.existsSync(packagePath)) {
    throw new Error(`SPFx package not found: ${packagePath}`);
  }

  const packageBuffer = fs.readFileSync(packagePath);
  const zip = await JSZip.loadAsync(packageBuffer);
  const relsFile = zip.file(relsPath);
  if (!relsFile) {
    throw new Error(`SPFx package does not contain ${relsPath}`);
  }

  const relsXml = await relsFile.async('string');
  zip.file(relsPath, addClientAssetRelationship(relsXml));
  zip.file(clientAssetPath, fs.readFileSync(iconSourcePath));

  const updatedPackage = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE'
  });
  fs.writeFileSync(packagePath, updatedPackage);
}

async function main() {
  ensureIconExists();
  patchDebugPackage();
  await patchSppkg();
  console.log(`Ensured catalog icon client asset: ${clientAssetPath}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
