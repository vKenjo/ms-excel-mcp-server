#!/usr/bin/env python3
"""
Test script to verify the new Excel MCP server tools are registered and available.

This script starts the MCP server and checks if the new tools are listed.
"""

import subprocess
import json
import sys
import time
import signal
import os


def test_tool_registration():
    """Test that all new tools are properly registered"""

    # Start the MCP server process
    exe_path = os.path.join(os.path.dirname(__file__), "..", "excel-mcp-server.exe")
    if not os.path.exists(exe_path):
        exe_path = "./excel-mcp-server.exe"

    try:
        # Start server process
        proc = subprocess.Popen(
            [exe_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Send initialize request
        initialize_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0.0"},
            },
        }

        proc.stdin.write(json.dumps(initialize_request) + "\n")
        proc.stdin.flush()

        # Read response
        response_line = proc.stdout.readline()
        if response_line:
            response = json.loads(response_line.strip())
            print("Initialize response received:", json.dumps(response, indent=2))

        # Send tools/list request
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {},
        }

        proc.stdin.write(json.dumps(tools_request) + "\n")
        proc.stdin.flush()

        # Read tools response
        tools_response_line = proc.stdout.readline()
        if tools_response_line:
            tools_response = json.loads(tools_response_line.strip())
            print("Tools list response:", json.dumps(tools_response, indent=2))

            # Check for our new tools
            expected_tools = [
                "excel_add_data_validation",
                "excel_add_conditional_formatting",
                "excel_execute_vba",
                "excel_add_vba_module",
            ]

            if "result" in tools_response and "tools" in tools_response["result"]:
                available_tools = [
                    tool["name"] for tool in tools_response["result"]["tools"]
                ]
                print(f"\nAvailable tools: {available_tools}")

                missing_tools = []
                for tool in expected_tools:
                    if tool in available_tools:
                        print(f"✅ {tool} - FOUND")
                    else:
                        print(f"❌ {tool} - MISSING")
                        missing_tools.append(tool)

                if not missing_tools:
                    print("\n🎉 All new tools are successfully registered!")
                    return True
                else:
                    print(f"\n❌ Missing tools: {missing_tools}")
                    return False
            else:
                print("❌ No tools found in response")
                return False

    except Exception as e:
        print(f"Error testing tools: {e}")
        return False
    finally:
        # Cleanup process
        try:
            proc.terminate()
            proc.wait(timeout=5)
        except:
            proc.kill()

    return False


if __name__ == "__main__":
    print("Testing Excel MCP Server - New Tools Registration")
    print("=" * 50)

    success = test_tool_registration()

    if success:
        print("\n✅ Test PASSED - All new tools are available!")
        sys.exit(0)
    else:
        print("\n❌ Test FAILED - Some tools are missing!")
        sys.exit(1)
