const { spawn } = require('child_process');
const path = require('path');
const { series, task } = require('gulp');

const heftScript = path.join(
  __dirname,
  'node_modules',
  '@rushstack',
  'heft',
  'bin',
  'heft'
);

function runHeft(args, done) {
  const heft = spawn(process.execPath, [heftScript, ...args], { stdio: 'inherit' });

  heft.on('error', done);
  heft.on('close', (code) => {
    done(code === 0 ? undefined : new Error(`Heft exited with code ${code}`));
  });
}

task('build', function build(done) {
  runHeft(['test', '--clean', '--production'], (error) => {
    if (error) {
      done(error);
      return;
    }

    runHeft(['package-solution', '--production'], done);
  });
});

task('serve', function serve(done) {
  runHeft(['start', '--clean'], done);
});

task('default', series('build'));
