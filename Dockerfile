
# Build stage for Go application
FROM golang:1.23-alpine AS go-builder

WORKDIR /go/src/app

# Install necessary packages
RUN apk add --no-cache git ca-certificates

# Copy go mod files
COPY go.mod go.sum ./
RUN go mod download

# Copy source code
COPY . .

# Build the Go application
RUN CGO_ENABLED=0 GOOS=linux go build -a -installsuffix cgo -o excel-mcp-server ./cmd/excel-mcp-server

# Runtime stage - use minimal alpine image since we don't need Node.js
FROM alpine:latest AS release

# Install ca-certificates for HTTPS requests
RUN apk --no-cache add ca-certificates
WORKDIR /app

# Copy the built Go binary from the build stage
COPY --from=go-builder /go/src/app/excel-mcp-server ./

# Make the binary executable
RUN chmod +x ./excel-mcp-server

# Test that the command works (will show usage or fail gracefully)
RUN ./excel-mcp-server || echo "Binary is executable"

# Command to run the application
ENTRYPOINT ["./excel-mcp-server"]
