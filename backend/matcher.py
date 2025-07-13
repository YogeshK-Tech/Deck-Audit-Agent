import openai
from fuzzywuzzy import fuzz
import re
from typing import List, Dict, Any, Optional
import os
import json
import math

class NumberMatcher:
    def __init__(self):
        # Set your OpenAI API key here or in environment
        openai.api_key = os.getenv("OPENAI_API_KEY", "your-api-key-here")
        self.tolerance = 0.05  # 5% tolerance for number matching
        
    def match_numbers(self, ppt_data: List[Dict], excel_data: Dict) -> List[Dict[str, Any]]:
        """Match numbers from PPT with Excel data"""
        try:
            audit_results = []
            
            # Flatten all numbers from all slides
            all_ppt_numbers = []
            for slide in ppt_data:
                for number in slide["numbers"]:
                    all_ppt_numbers.append(number)
            
            # Flatten all numbers from all Excel sheets
            all_excel_numbers = []
            for file_data in excel_data.values():
                for sheet_data in file_data["sheets"].values():
                    for number in sheet_data["numbers"]:
                        number["source_file"] = file_data["filename"]
                        all_excel_numbers.append(number)
            
            print(f"Matching {len(all_ppt_numbers)} PPT numbers against {len(all_excel_numbers)} Excel numbers")
            
            # Match each PPT number
            for ppt_number in all_ppt_numbers:
                match_result = self._find_best_match(ppt_number, all_excel_numbers)
                audit_results.append(match_result)
            
            return audit_results
            
        except Exception as e:
            print(f"Error in matching process: {str(e)}")
            raise Exception(f"Matching failed: {str(e)}")
    
    def _find_best_match(self, ppt_number: Dict, excel_numbers: List[Dict]) -> Dict[str, Any]:
        """Find the best match for a PPT number in Excel data"""
        try:
            best_match = None
            best_score = 0
            
            # First, try exact numerical match
            for excel_num in excel_numbers:
                if self._numbers_match(ppt_number["parsed_value"], excel_num["value"]):
                    # Check context similarity
                    context_score = self._calculate_context_similarity(ppt_number, excel_num)
                    total_score = 0.7 + (0.3 * context_score / 100)  # 70% number, 30% context
                    
                    if total_score > best_score:
                        best_score = total_score
                        best_match = excel_num
            
            # If no exact match, try fuzzy matching with context
            if best_match is None:
                for excel_num in excel_numbers:
                    context_score = self._calculate_context_similarity(ppt_number, excel_num)
                    
                    # Lower threshold for context-based matching
                    if context_score > 60:
                        total_score = context_score / 100
                        if total_score > best_score:
                            best_score = total_score
                            best_match = excel_num
            
            # Determine status and create result
            if best_match is None:
                return self._create_untraceable_result(ppt_number)
            elif self._numbers_match(ppt_number["parsed_value"], best_match["value"]):
                return self._create_match_result(ppt_number, best_match)
            else:
                return self._create_mismatch_result(ppt_number, best_match)
                
        except Exception as e:
            print(f"Error finding match for number: {str(e)}")
            return self._create_error_result(ppt_number, str(e))
    
    def _numbers_match(self, ppt_value: float, excel_value: float) -> bool:
        """Check if two numbers match within tolerance"""
        if ppt_value == 0 and excel_value == 0:
            return True
        if ppt_value == 0 or excel_value == 0:
            return False
        
        # Calculate relative difference
        diff = abs(ppt_value - excel_value) / max(abs(ppt_value), abs(excel_value))
        return diff <= self.tolerance
    
    def _calculate_context_similarity(self, ppt_number: Dict, excel_number: Dict) -> float:
        """Calculate similarity between contexts using fuzzy matching"""
        ppt_context = f"{ppt_number.get('context', '')} {ppt_number.get('raw_text', '')}"
        excel_context = f"{excel_number.get('context', '')} {excel_number.get('original_text', '')}"
        
        # Clean and normalize contexts
        ppt_context = self._clean_context(ppt_context)
        excel_context = self._clean_context(excel_context)
        
        if not ppt_context or not excel_context:
            return 0
        
        # Use fuzzy matching
        similarity = fuzz.partial_ratio(ppt_context.lower(), excel_context.lower())
        return similarity
    
    def _clean_context(self, context: str) -> str:
        """Clean and normalize context text"""
        # Remove special characters and extra spaces
        context = re.sub(r'[^\w\s]', ' ', context)
        context = re.sub(r'\s+', ' ', context)
        return context.strip()
    
    def _create_match_result(self, ppt_number: Dict, excel_match: Dict) -> Dict[str, Any]:
        """Create result for successful match"""
        return {
            "slide": ppt_number["slide_number"],
            "text": ppt_number["raw_text"],
            "status": "Match",
            "ppt_value": ppt_number["parsed_value"],
            "excel_value": excel_match["value"],
            "suggested_fix": None,
            "excel_sheet": excel_match["sheet_name"],
            "excel_file": excel_match.get("source_file", ""),
            "cell": excel_match["cell_reference"],
            "reasoning": f"Values match within tolerance. PPT: {ppt_number['parsed_value']}, Excel: {excel_match['value']}",
            "context": ppt_number.get("context", ""),
            "confidence": 0.95
        }
    
    def _create_mismatch_result(self, ppt_number: Dict, excel_match: Dict) -> Dict[str, Any]:
        """Create result for mismatched numbers"""
        # Format suggested fix based on original format
        suggested_fix = self._format_suggestion(excel_match["value"], ppt_number["type"])
        
        return {
            "slide": ppt_number["slide_number"],
            "text": ppt_number["raw_text"],
            "status": "Mismatch",
            "ppt_value": ppt_number["parsed_value"],
            "excel_value": excel_match["value"],
            "suggested_fix": suggested_fix,
            "excel_sheet": excel_match["sheet_name"],
            "excel_file": excel_match.get("source_file", ""),
            "cell": excel_match["cell_reference"],
            "reasoning": f"Number mismatch detected. PPT shows {ppt_number['parsed_value']}, but Excel shows {excel_match['value']}",
            "context": ppt_number.get("context", ""),
            "confidence": 0.80
        }
    
    def _create_untraceable_result(self, ppt_number: Dict) -> Dict[str, Any]:
        """Create result for untraceable numbers"""
        return {
            "slide": ppt_number["slide_number"],
            "text": ppt_number["raw_text"],
            "status": "Untraceable",
            "ppt_value": ppt_number["parsed_value"],
            "excel_value": None,
            "suggested_fix": None,
            "excel_sheet": None,
            "excel_file": None,
            "cell": None,
            "reasoning": "No matching data found in Excel sheets",
            "context": ppt_number.get("context", ""),
            "confidence": 0.0
        }
    
    def _create_error_result(self, ppt_number: Dict, error_msg: str) -> Dict[str, Any]:
        """Create result for errors during matching"""
        return {
            "slide": ppt_number["slide_number"],
            "text": ppt_number["raw_text"],
            "status": "Error",
            "ppt_value": ppt_number["parsed_value"],
            "excel_value": None,
            "suggested_fix": None,
            "excel_sheet": None,
            "excel_file": None,
            "cell": None,
            "reasoning": f"Error during matching: {error_msg}",
            "context": ppt_number.get("context", ""),
            "confidence": 0.0
        }
    
    def _format_suggestion(self, value: float, number_type: str) -> str:
        """Format suggested fix based on original number type"""
        try:
            if number_type == "percentage":
                return f"{value:.1f}%"
            elif number_type == "currency":
                if value >= 10000000:  # Crores
                    return f"₹{value/10000000:.1f} Cr"
                elif value >= 100000:  # Lakhs
                    return f"₹{value/100000:.1f} Lac"
                elif value >= 1000:  # Thousands
                    return f"₹{value/1000:.1f}K"
                else:
                    return f"₹{value:.0f}"
            else:
                if value >= 10000000:
                    return f"{value/10000000:.1f} Cr"
                elif value >= 100000:
                    return f"{value/100000:.1f} Lac"
                elif value >= 1000:
                    return f"{value/1000:.1f}K"
                else:
                    return f"{value:.0f}"
        except:
            return str(value)