import fs from 'node:fs/promises';
import path from 'node:path';
const folder=path.resolve('dist');
let html=await fs.readFile(path.join(folder,'index.html'),'utf8');
const css=html.match(/<link[^>]+href="([^"]+\.css)"[^>]*>/);
const js=html.match(/<script[^>]+src="([^"]+\.js)"[^>]*><\/script>/);
if(!css||!js)throw new Error('Expected a Vite HTML entry with one CSS and one JS file.');
const read=async file=>fs.readFile(path.join(folder,file.replace(/^\.\//,'')),'utf8');
let script=await read(js[1]);
const imports=[...script.matchAll(/import\("(\.\/[^"\n]+\.js)"\)/g)];
for(const match of imports){
 const target=path.posix.join(path.posix.dirname(js[1]),match[1]);
 const module=await read(target);
 if(/(?:from\s*|import\s*\()["']\.\//.test(module))throw new Error('Portable module has unresolved relative imports.');
 script=script.replace(match[0],`import("data:text/javascript;base64,${Buffer.from(module).toString('base64')}")`);
}
const style=await read(css[1]);html=html.replace(css[0],()=>`<style>${style}</style>`);
html=html.replace(js[0],()=>`<script type="module">${script.replace(/<\/script/gi,'<\\/script')}</script>`);
const destination=path.join(folder,'REI-Print-Assistant.html');await fs.writeFile(destination,html);
console.log(`Portable local-only app: ${destination}`);
