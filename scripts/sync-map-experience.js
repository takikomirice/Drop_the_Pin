const fs=require('node:fs');
const path=require('node:path');
const esbuild=require('esbuild');
const root=path.resolve(__dirname,'..');
const check=process.argv.includes('--check');
const built=esbuild.buildSync({entryPoints:[path.join(root,'src/map-experience.js')],bundle:true,format:'iife',globalName:'DtpMapExperience',minify:true,target:'es2020',write:false}).outputFiles[0].text;
const plugin=fs.readFileSync(path.join(root,'node_modules/leaflet.markercluster/dist/leaflet.markercluster.js'),'utf8');
const css=fs.readFileSync(path.join(root,'src/map-experience.css'),'utf8');
const region='<!-- MAP_EXPERIENCE_START -->\n<style>\n'+css+'\n</style>\n<script>\n'+(plugin+'\n'+built).replace(/<\/script/gi,'<\\/script')+'\n</script>\n<!-- MAP_EXPERIENCE_END -->';
let changed=false;
function write(file,content){const p=path.join(root,file);if(fs.existsSync(p)&&fs.readFileSync(p,'utf8')===content)return;changed=true;if(!check)fs.writeFileSync(p,content);else console.error('Out of date: '+file);}
for(const file of ['index.html','shared.html']) {
  const source=fs.readFileSync(path.join(root,file),'utf8');
  const pattern=/<!-- MAP_EXPERIENCE_START -->[\s\S]*?<!-- MAP_EXPERIENCE_END -->/;
  const next=pattern.test(source)?source.replace(pattern,()=>region):source.replace('</head>',()=>region+'\n</head>');
  write(file,next);
}
for(const [pkg,name] of [['leaflet.markercluster','MIT-LICENCE.txt']]) {
  write('vendor/'+pkg+'-LICENSE.txt',fs.readFileSync(path.join(root,'node_modules',pkg,name),'utf8'));
}
if(check&&changed)process.exitCode=1;
