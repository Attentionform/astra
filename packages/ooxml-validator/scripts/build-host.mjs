import { execSync } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const packageRoot = path.resolve(__dirname, '..');

const platform = process.platform;
const arch = process.arch;

let rid = '';
if (platform === 'linux' && arch === 'x64') rid = 'linux-x64';
else if (platform === 'linux' && arch === 'arm64') rid = 'linux-arm64';
else if (platform === 'darwin' && arch === 'x64') rid = 'osx-x64';
else if (platform === 'darwin' && arch === 'arm64') rid = 'osx-arm64';
else if (platform === 'win32' && arch === 'x64') rid = 'win-x64';
else if (platform === 'win32' && arch === 'arm64') rid = 'win-arm64';
else {
  console.error(`Unsupported platform for native build: ${platform}-${arch}`);
  process.exit(1);
}

const outDir = path.join(packageRoot, 'native', 'out', rid);
const csproj = path.join(packageRoot, 'native', 'ooxml-validator.csproj');

const cmd = `dotnet publish "${csproj}" -c Release -r ${rid} --self-contained true -o "${outDir}" /p:PublishSingleFile=true /p:InvariantGlobalization=true`;
console.log(`[build:native] Building native binary for ${rid} into ${outDir}`);
execSync(cmd, { stdio: 'inherit' });
