import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const port = process.env.PORT || 3002;

const serve = spawn('serve', ['-s', 'build', '-l', port], {
  stdio: 'inherit',
  cwd: join(__dirname, 'build')
});

serve.on('close', (code) => {
  console.log(`child process exited with code ${code}`);
});

console.log(`Server running on port ${port}`);
