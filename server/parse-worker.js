import fs from 'node:fs/promises';
const [path, extension] = process.argv.slice(2);
try {
  const buffer = await fs.readFile(path);
  let text;
  if (extension === '.docx') {
    const { default: mammoth } = await import('mammoth');
    text = (await mammoth.extractRawText({ buffer })).value;
  } else if (extension === '.pdf') {
    const { PDFParse } = await import('pdf-parse');
    const parser = new PDFParse({ data: buffer });
    try { text = (await parser.getText()).text; } finally { await parser.destroy(); }
  } else if (extension === '.txt') text = buffer.toString('utf8');
  else throw new Error('Unsupported document');
  if (Buffer.byteLength(text) > 2 * 1024 * 1024) throw new Error('Text limit');
  process.stdout.write(text);
} catch { process.exitCode = 1; }
