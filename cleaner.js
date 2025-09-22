import fs from 'fs';
import path from 'path';
import { Document, Packer, Paragraph, Table, TextRun } from 'docx'; // For reading, limited support
import readdirp from 'readdirp';

// Map of Italian accented characters to replacements
const charMap = {
  'à': 'a', 'À': 'A',
  'è': 'e', 'È': 'E',
  'é': 'e', 'É': 'E',
  'ì': 'i', 'Ì': 'I',
  'ò': 'o', 'Ò': 'O',
  'ù': 'u', 'Ù': 'U',
  '—': ', '
};

const exceptionalWords = new Set(['μ', 'µ']);

// Replace unwanted characters except ASCII, bullet (•), and comma (,)
function removeUnwantedSpecials(text) {
  let removedChars = [];
  let newText = '';
  for (const ch of text) {
    if (ch.charCodeAt(0) <= 127 || ch === '•' || ch === ',') {
      newText += ch;
    } else {
      removedChars.push(ch);
    }
  }
  return { newText, removedChars };
}

function cleanText(text, location = '', tableMode = false) {
  if (exceptionalWords.has(text.trim())) {
    return text;
  }

  const changes = [];

  if (text.includes('\u00A0')) {
    changes.push('non-breaking space');
    text = text.replace(/\u00A0/g, ' ');
  }

  for (const [k, v] of Object.entries(charMap)) {
    if (text.includes(k)) {
      changes.push(k);
      text = text.split(k).join(v);
    }
  }

  const { newText, removedChars } = removeUnwantedSpecials(text);
  if (removedChars.length) {
    changes.push(...removedChars);
  }
  text = newText;

  if (changes.length) {
    console.log(`${location}: ${[...new Set(changes)].sort().join(', ')}`);
  }

  return text;
}

async function processDocxFile(filePath) {
  console.log(`Processing file: ${path.basename(filePath)}`);

  // NOTE: docx npm package has limited reading support — for full editing use another library or
  // convert to plain text, edit, then recreate the docx

  // For demo, we'll read file buffer and create a new doc with cleaned paragraphs (losing formatting)

  // You might want to extract paragraphs with a library like mammoth
  // Here's a simplified way using mammoth:

  import mammoth from 'mammoth';
  const { value: text } = await mammoth.extractRawText({ path: filePath });

  // Split text into paragraphs roughly
  let paragraphs = text.split('\n').filter(p => p.trim() !== '');

  // Clean paragraphs
  paragraphs = paragraphs.map((p, i) => cleanText(p, `Paragraph ${i + 1}`));

  // Create new doc with cleaned paragraphs (basic, no tables)
  const doc = new Document({
    sections: [{
      children: paragraphs.map(p => new Paragraph(p)),
    }],
  });

  const buffer = await Packer.toBuffer(doc);

  await fs.promises.writeFile(filePath, buffer);

  console.log(`Overwritten: ${filePath}\n${'-'.repeat(50)}`);
}

async function main() {
  const folderPath = "C:\\Users\\Madhan\\Write Final";

  const files = await fs.promises.readdir(folderPath);

  for (const fileName of files) {
    if (fileName.toLowerCase().endsWith('.docx')) {
      const filePath = path.join(folderPath, fileName);
      await processDocxFile(filePath);
    }
  }

  console.log("All files processed.");
}

main().catch(console.error);
