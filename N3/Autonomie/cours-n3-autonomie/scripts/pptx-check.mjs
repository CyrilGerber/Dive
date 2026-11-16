import fs from "node:fs";
import path from "node:path";
import process from "node:process";

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const manifest = path.join(root, "manifest.yml");
const slides = path.join(root, "slides");

if (!fs.existsSync(manifest) || !fs.existsSync(slides)) {
  console.error("Projet incomplet : manifest.yml ou slides/ absent.");
  process.exit(1);
}

const slideFiles = fs.readdirSync(slides).filter((name) => name.endsWith(".md"));
console.log(`Précontrôle réussi : ${slideFiles.length} fichiers Markdown détectés.`);
console.log("La génération PowerPoint n’est volontairement pas implémentée à ce stade.");
