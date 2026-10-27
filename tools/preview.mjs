import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const [source, target, range='A1:G4'] = process.argv.slice(2);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const preview = await workbook.render({sheetName:workbook.worksheets.getItemAt(0).name,range,scale:1,format:'png'});
await fs.writeFile(target,new Uint8Array(await preview.arrayBuffer()));
console.log(JSON.stringify({preview:target,range}));
