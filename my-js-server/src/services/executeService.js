// src/services/executeService.js


exports.runCommand = async (param1, param2) => {
  return new Promise((resolve, reject) => {
    // For example, using some shell command
    const command = `echo ${param1} and ${param2}`;

    exec(command, (error, stdout, stderr) => {
      if (error) {
        console.error('Exec error:', error);
        reject(error);
      }
      resolve(stdout.trim());
    });
  });
};
