module.exports = {
  apps : [{
    name: "backend",
    cwd: "./backend",
    script: "npm",
    args: "start",
    env: {
      NODE_ENV: "production",
      PORT: 5001
    },
  }, {
    name: "frontend",
    cwd: "./frontend",
    script: "npm",
    args: "start",
    env: {
      NODE_ENV: "production",
      PORT: 3002
    },
  }]
}
