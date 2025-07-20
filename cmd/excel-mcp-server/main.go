package main

import (
	"fmt"
	"log"
	"os"

	"github.com/vKenjo/ms-excel-mcp-server/internal/server"
)

var (
	version = "dev"
)

func main() {
	// Add panic recovery
	defer func() {
		if r := recover(); r != nil {
			log.Printf("Fatal panic: %v", r)
			os.Exit(1)
		}
	}()
	
	s := server.New(version)
	err := s.Start()
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to start the server: %v\n", err)
		os.Exit(1)
	}
}
