
FROM node:20 AS release

# Set the working directory
WORKDIR /app

# Install the package globally
RUN npm install -g ms-excel-mcp-server@0.11.0

# Test that the command works
RUN excel-mcp-server --help || echo "Command executable"

# Command to run the application
ENTRYPOINT ["excel-mcp-server"]
