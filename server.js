// Root-level entry point for production deployments (e.g., Hostinger Node.js App)
// Sets the environment to production and loads the bundled server
process.env.NODE_ENV = "production";

// Dynamically import the compiled CommonJS server bundle
import("./dist/server.cjs");
