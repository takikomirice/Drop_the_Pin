const { defineConfig, devices } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './tests/browser',
  timeout: 60_000,
  fullyParallel: false,
  workers: 1,
  reporter: 'line',
  webServer: {
    command: 'node tests/browser/harness-server.js',
    url: 'http://127.0.0.1:4173/__health',
    reuseExistingServer: false,
    timeout: 30_000
  },
  use: {
    baseURL: 'http://127.0.0.1:4173',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure'
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'], browserName: 'chromium' }
    }
  ]
});
