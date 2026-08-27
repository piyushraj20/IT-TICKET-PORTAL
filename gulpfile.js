const { spawn } = require('child_process');
const path = require('path');
const { task } = require('gulp');

const heftCommand = path.join(
  __dirname,
  'node_modules',
  '.bin',
  process.platform === 'win32' ? 'heft.cmd' : 'heft'
);

task('serve', function serve(done) {
  const server = spawn(heftCommand, ['start', '--clean'], {
    stdio: 'inherit',
    // Windows .cmd files must be launched through a shell.
    shell: process.platform === 'win32'
  });

  server.on('error', done);
  server.on('close', (code) => {
    if (code === 0) {
      done();
    } else {
      done(new Error(`Heft server exited with code ${code}`));
    }
  });
});
