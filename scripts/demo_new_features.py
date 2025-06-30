#!/usr/bin/env python3
"""
Excel MCP Server - New Features Demo Script

This script demonstrates the new data validation, conditional formatting, and VBA features
added to the Excel MCP server.

Usage:
    python demo_new_features.py

Requirements:
    - Excel MCP Server running
    - Python with requests library
    - Excel file for testing
"""

import json
import sys
import os
from pathlib import Path

# Example usage functions for the new Excel tools


def create_data_validation_dropdown_example():
    """Example: Create a dropdown list validation"""
    return {
        "tool": "excel_add_data_validation",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "cellRange": "A1:A10",
            "validationType": "list",
            "options": {
                "dropdownList": ["High", "Medium", "Low"],
                "showInputMessage": True,
                "inputTitle": "Priority Level",
                "inputMessage": "Please select a priority level from the dropdown",
                "showErrorMessage": True,
                "errorTitle": "Invalid Priority",
                "errorMessage": "Please select High, Medium, or Low",
            },
        },
    }


def create_data_validation_number_range_example():
    """Example: Create number range validation"""
    return {
        "tool": "excel_add_data_validation",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "cellRange": "B1:B10",
            "validationType": "whole",
            "options": {
                "operator": "between",
                "formula1": "1",
                "formula2": "100",
                "showInputMessage": True,
                "inputTitle": "Enter Score",
                "inputMessage": "Please enter a score between 1 and 100",
                "showErrorMessage": True,
                "errorTitle": "Invalid Score",
                "errorMessage": "Score must be between 1 and 100",
            },
        },
    }


def create_conditional_formatting_highlight_example():
    """Example: Highlight cells greater than 50"""
    return {
        "tool": "excel_add_conditional_formatting",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "cellRange": "C1:C10",
            "conditions": {
                "type": "cellValue",
                "criteria": "greaterThan",
                "value1": "50",
                "format": {
                    "font": {"bold": True, "color": "#FFFFFF"},
                    "fill": {"type": "pattern", "color": ["#FF0000"]},
                },
            },
        },
    }


def create_conditional_formatting_color_scale_example():
    """Example: Color scale formatting"""
    return {
        "tool": "excel_add_conditional_formatting",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "cellRange": "D1:D10",
            "conditions": {
                "type": "colorScale",
                "colorScale": {
                    "minType": "min",
                    "minColor": "#FF0000",
                    "maxType": "max",
                    "maxColor": "#00FF00",
                },
            },
        },
    }


def create_conditional_formatting_data_bars_example():
    """Example: Data bars formatting"""
    return {
        "tool": "excel_add_conditional_formatting",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "cellRange": "E1:E10",
            "conditions": {
                "type": "dataBar",
                "dataBar": {
                    "minType": "min",
                    "maxType": "max",
                    "color": "#0000FF",
                    "showValue": True,
                },
            },
        },
    }


def create_vba_execution_example():
    """Example: Execute VBA code (Windows OLE only)"""
    return {
        "tool": "excel_execute_vba",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "vbaCode": 'Range("F1").Value = "Hello from VBA!"',
        },
    }


def create_vba_module_example():
    """Example: Add VBA module (Windows OLE only)"""
    vba_code = """Sub CalculateTotal()
    Dim i As Integer
    Dim total As Double
    total = 0
    
    For i = 1 To 10
        total = total + Range("C" & i).Value
    Next i
    
    Range("C11").Value = total
    Range("C11").Font.Bold = True
End Sub

Function MultiplyByTwo(value As Double) As Double
    MultiplyByTwo = value * 2
End Function"""

    return {
        "tool": "excel_add_vba_module",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "moduleName": "CalculationModule",
            "vbaCode": vba_code,
        },
    }


def print_examples():
    """Print all examples as JSON for easy copy-paste"""
    examples = [
        ("Data Validation - Dropdown List", create_data_validation_dropdown_example()),
        (
            "Data Validation - Number Range",
            create_data_validation_number_range_example(),
        ),
        (
            "Conditional Formatting - Highlight",
            create_conditional_formatting_highlight_example(),
        ),
        (
            "Conditional Formatting - Color Scale",
            create_conditional_formatting_color_scale_example(),
        ),
        (
            "Conditional Formatting - Data Bars",
            create_conditional_formatting_data_bars_example(),
        ),
        ("VBA Execution", create_vba_execution_example()),
        ("VBA Module Addition", create_vba_module_example()),
    ]

    print("Excel MCP Server - New Features Examples")
    print("=" * 50)
    print()

    for title, example in examples:
        print(f"## {title}")
        print()
        print("```json")
        print(json.dumps(example, indent=2))
        print("```")
        print()


def create_sample_excel_data():
    """Example: Create sample data in Excel for testing"""
    return {
        "tool": "excel_write_to_sheet",
        "arguments": {
            "fileAbsolutePath": "C:\\temp\\example.xlsx",
            "sheetName": "Sheet1",
            "newSheet": False,
            "range": "C1:C10",
            "values": [
                ["25"],
                ["67"],
                ["43"],
                ["89"],
                ["12"],
                ["76"],
                ["34"],
                ["91"],
                ["55"],
                ["38"],
            ],
        },
    }


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--sample-data":
        print("Sample Data Creation Example:")
        print(json.dumps(create_sample_excel_data(), indent=2))
    else:
        print_examples()
