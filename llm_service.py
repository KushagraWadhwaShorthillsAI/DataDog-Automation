#!/usr/bin/env python3
"""
Modular LLM Service for Error Categorization
Supports multiple LLM providers: Azure OpenAI, Google Gemini, OpenAI, etc.
"""

import os
import json
import time
from typing import Dict, Optional, Any
from dotenv import load_dotenv
from abc import ABC, abstractmethod

# Load environment variables
load_dotenv()

class LLMProvider(ABC):
    """Abstract base class for LLM providers"""
    
    @abstractmethod
    def categorize_error(self, error_message: str) -> Dict[str, Any]:
        """Categorize an error message and return structured result"""
        pass

class AzureOpenAIProvider(LLMProvider):
    """Azure OpenAI provider implementation"""
    
    def __init__(self):
        self.api_key = os.getenv('AZURE_OPENAI_API_KEY')
        self.api_version = os.getenv('AZURE_OPENAI_API_VERSION', '2023-05-15')
        self.endpoint = os.getenv('AZURE_OPENAI_ENDPOINT')
        self.deployment = os.getenv('AZURE_OPENAI_DEPLOYMENT', 'gpt-4')
        
        if not all([self.api_key, self.endpoint, self.deployment]):
            raise ValueError("Missing required Azure OpenAI configuration in .env file")
        
        # Import Azure OpenAI
        try:
            from openai import AzureOpenAI
            self.client = AzureOpenAI(
                api_key=self.api_key,
                api_version=self.api_version,
                azure_endpoint=self.endpoint
            )
        except ImportError:
            raise ImportError("Please install openai package: pip install openai")
    
    def categorize_error(self, error_message: str) -> Dict[str, Any]:
        """Categorize error using Azure OpenAI"""
        try:
            prompt = self._build_prompt(error_message)
            
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": "You are an expert error categorization system."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=500
            )
            
            response_text = response.choices[0].message.content.strip()
            return self._parse_response(response_text)
            
        except Exception as e:
            print(f"⚠️ Azure OpenAI categorization failed: {e}")
            return self._get_fallback_result(error_message)
    
    def _build_prompt(self, error_message: str) -> str:
        """Build the categorization prompt"""
        return f"""
You are an Expert Error Analysis Engine. Your task is to analyze web application error messages with high precision, providing a structured, multi-faceted categorization for deeper insights.

## 1. CATEGORY DEFINITIONS

1. Timeout Errors
- Description: Any error indicating that an operation did not complete within an expected timeframe.
- Sub-Categories: Request Timeout, Connection Timeout, Operation Timeout, Gateway Timeout.
- Examples: "Request timeout", "Connection timed out", "deadline exceeded", "504 Gateway Time-out"

2. Network/Connection Errors
- Description: Failures related to network connectivity, sockets, or the inability to establish a connection. Distinct from timeouts.
- Sub-Categories: Connection Refused, Connection Aborted, Host Not Found, Network Unreachable, Socket Error.
- Examples: "Connection Failed", "Network unreachable", "Connection aborted", "ECONNRESET", "Remote end closed without response"

3. Authentication/Authorization Errors
- Description: Errors related to user identity verification (authentication) or permission levels (authorization).
- Sub-Categories: Authentication Failed, Unauthorized Access, Permission Denied, Forbidden, Invalid Credentials.
- Examples: "Unauthorized", "Permission denied", "Authentication failed", "403 Forbidden", "Invalid API Key"

4. Resource Not Found Errors
- Description: Errors indicating that a requested resource, asset, or document could not be located.
- Sub-Categories: 404 Not Found, Document Not Found, File Not Found, No Results.
- Examples: "Not found", "No document selected", "Resource not found", "Contains no results"

5. Data Validation/Payload Errors
- Description: Errors caused by malformed, incorrect, or incomplete data sent by the client.
- Sub-Categories: Validation Failed, Bad Request, Invalid Payload, Missing Field.
- Examples: "Invalid data payload", "Validation failed", "400 Bad request", "Missing required field 'user_id'"

6. Internal Server Errors
- Description: General, non-specific server-side errors indicating a problem with the server's execution, but not a specific application-level exception. Often represented by 5xx status codes.
- Sub-Categories: 500 Internal Server Error, Server Overloaded, Bad Gateway.
- Examples: "Internal server error", "Server error", "500 error", "502 Bad Gateway"

7. LLM Service Errors
- Description: Errors originating specifically from a Large Language Model (LLM) service or library (e.g., OpenAI, Anthropic, LiteLLM).
- Sub-Categories: API Error, Rate Limit Error, Context Window Exceeded, Service Unavailable, Quota Exceeded.
- Examples: "litellm.ServiceUnavailableError", "ContextWindowExceededError", "RateLimitError", "Token length exceeds", "The model is currently overloaded"

8. Query/Parameter Errors
- Description: Errors related to the structure, syntax, or values of query parameters in a request.
- Sub-Categories: Invalid Query, Missing Parameter, Invalid Filter Type.
- Examples: "Missing filterType", "Invalid query", "Parameter 'sort_by' is not valid"

9. Application Exception Errors
- Description: Specific, unhandled exceptions originating from the application's code logic (e.g., Python, Node.js). Distinct from general Internal Server Errors.
- Sub-Categories: TypeError, AttributeError, NullPointerException, KeyError, ValueError.
- Examples: "TypeError: 'NoneType' object is not iterable", "AttributeError: 'object' has no attribute 'user'", "KeyError: 'config'"

10. Service Configuration Errors
- Description: Errors related to application setup, model mapping, or failure to fetch necessary configuration.
- Sub-Categories: Model Mapping Unavailable, Configuration Fetch Failed, Invalid Setup.
- Examples: "Model configuration unavailable", "Failed to fetch model mapping"

11. Data Format Errors
- Description: Errors that occur while parsing or processing data that does not conform to the expected format.
- Sub-Categories: JSON Parse Error, XML Parse Error, Invalid Data Structure.
- Examples: "JSON parse error: Unexpected token", "Invalid format", "Data structure mismatch"

12. Streaming Errors
- Description: Failures that occur during an active data stream.
- Sub-Categories: Stream Interrupted, Streaming Failed, Incomplete Stream.
- Examples: "Error raised while streaming", "Stream interrupted", "Streaming failed"

13. Request/Response Logging Errors
- Description: JSON objects or structured data containing request metadata, session information, or logging data rather than actual error messages.
- Sub-Categories: Request Metadata, Session Data, Logging Data, Response Data.
- Examples: JSON objects starting with {{"RequestId":, {{"session_id":, logging data structures

14. Feature Configuration Errors
- Description: Errors related to feature flags, configuration settings, or application feature management.
- Sub-Categories: Feature Flag Error, Configuration Unavailable, Feature Disabled.
- Examples: "Feature flag error", "Configuration unavailable", "Feature not enabled"

15. Other/Uncategorized Errors
- Description: Errors that do not fit into any of the above categories or are too ambiguous to categorize accurately.
- Sub-Categories: Unknown Error, Ambiguous Error, Unclassified Error.
- Examples: Generic error messages without clear context

## 2. INPUT
ERROR MESSAGE:
{error_message}

OPTIONAL APPLICATION CONTEXT:
This error comes from a web application service that processes user requests and may interact with LLM services, databases, and external APIs.

## 3. ANALYSIS AND OUTPUT INSTRUCTIONS
Carefully analyze the ERROR MESSAGE and any provided APPLICATION CONTEXT.
Identify the most accurate Primary Category from the list above.
Determine a concise Sub-Category based on the specific keywords in the error message.
Assign a Confidence Score (0-100%) representing your certainty in the categorization.
Write a brief, one-sentence Rationale explaining why you chose the category. If the error is ambiguous, note the second possibility here.

Return the output ONLY in the following JSON format:
{{
  "PrimaryCategory": "...",
  "SubCategory": "...",
  "ConfidenceScore": ...,
  "Rationale": "..."
}}

## 4. CRITICAL RULES
Your output MUST be a single, valid JSON object.
Do not add any text, explanation, markdown formatting, or code blocks before or after the JSON output.
Do not wrap the JSON in ```json``` or any other formatting.
PrimaryCategory must be one of the exact category names from the list.
SubCategory should be a specific term derived directly from the error message.
Be precise and prioritize the most direct root cause of the error.
ConfidenceScore must be a number between 0 and 100.
Rationale must be a single, clear sentence.
Output ONLY the raw JSON object, nothing else."""

    def _parse_response(self, response_text: str) -> Dict[str, Any]:
        """Parse the LLM response and return structured data"""
        try:
            # Clean up response text in case it has markdown formatting
            if response_text.startswith('```json'):
                response_text = response_text[7:]  # Remove ```json
            if response_text.endswith('```'):
                response_text = response_text[:-3]  # Remove ```
            response_text = response_text.strip()
            
            result = json.loads(response_text)
            primary_category = result.get('PrimaryCategory', 'Other/Uncategorized Errors')
            
            # Validate the response is one of our expected categories
            valid_categories = [
                'Timeout Errors', 'Network/Connection Errors', 'Authentication/Authorization Errors',
                'Resource Not Found Errors', 'Data Validation/Payload Errors', 'Internal Server Errors',
                'LLM Service Errors', 'Query/Parameter Errors', 'Application Exception Errors',
                'Service Configuration Errors', 'Data Format Errors', 'Streaming Errors',
                'Request/Response Logging Errors', 'Feature Configuration Errors', 'Other/Uncategorized Errors'
            ]
            
            if primary_category in valid_categories:
                return {
                    'category': primary_category,
                    'sub_category': result.get('SubCategory', 'Unknown'),
                    'confidence': result.get('ConfidenceScore', 0),
                    'rationale': result.get('Rationale', 'No rationale provided')
                }
            else:
                print(f"⚠️ Azure OpenAI returned unexpected category: '{primary_category}'")
                return self._get_fallback_result("Invalid category returned")
                
        except json.JSONDecodeError as e:
            print(f"⚠️ Failed to parse Azure OpenAI JSON response: {e}")
            print(f"Raw response: {response_text}")
            return self._get_fallback_result("JSON parse error")
    
    def _get_fallback_result(self, error_message: str) -> Dict[str, Any]:
        """Get fallback result when LLM fails"""
        return {
            'category': 'Other/Uncategorized Errors',
            'sub_category': 'LLM Processing Error',
            'confidence': 0,
            'rationale': f'Failed to process with Azure OpenAI: {error_message[:100]}'
        }

class GeminiProvider(LLMProvider):
    """Google Gemini provider implementation"""
    
    def __init__(self):
        self.api_key = os.getenv('GEMINI_API_KEY')
        if not self.api_key:
            raise ValueError("Missing GEMINI_API_KEY in .env file")
        
        try:
            import google.generativeai as genai
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
        except ImportError:
            raise ImportError("Please install google-generativeai package: pip install google-generativeai")
    
    def categorize_error(self, error_message: str) -> Dict[str, Any]:
        """Categorize error using Google Gemini"""
        try:
            prompt = self._build_prompt(error_message)
            response = self.model.generate_content(prompt)
            response_text = response.text.strip()
            return self._parse_response(response_text)
            
        except Exception as e:
            print(f"⚠️ Gemini categorization failed: {e}")
            return self._get_fallback_result(error_message)
    
    def _build_prompt(self, error_message: str) -> str:
        """Build the categorization prompt (same as Azure OpenAI)"""
        # Use the same prompt as Azure OpenAI
        azure_provider = AzureOpenAIProvider.__new__(AzureOpenAIProvider)
        return azure_provider._build_prompt(error_message)
    
    def _parse_response(self, response_text: str) -> Dict[str, Any]:
        """Parse the Gemini response"""
        try:
            # Clean up response text in case it has markdown formatting
            if response_text.startswith('```json'):
                response_text = response_text[7:]  # Remove ```json
            if response_text.endswith('```'):
                response_text = response_text[:-3]  # Remove ```
            response_text = response_text.strip()
            
            result = json.loads(response_text)
            primary_category = result.get('PrimaryCategory', 'Other/Uncategorized Errors')
            
            # Validate the response is one of our expected categories
            valid_categories = [
                'Timeout Errors', 'Network/Connection Errors', 'Authentication/Authorization Errors',
                'Resource Not Found Errors', 'Data Validation/Payload Errors', 'Internal Server Errors',
                'LLM Service Errors', 'Query/Parameter Errors', 'Application Exception Errors',
                'Service Configuration Errors', 'Data Format Errors', 'Streaming Errors',
                'Request/Response Logging Errors', 'Feature Configuration Errors', 'Other/Uncategorized Errors'
            ]
            
            if primary_category in valid_categories:
                return {
                    'category': primary_category,
                    'sub_category': result.get('SubCategory', 'Unknown'),
                    'confidence': result.get('ConfidenceScore', 0),
                    'rationale': result.get('Rationale', 'No rationale provided')
                }
            else:
                print(f"⚠️ Gemini returned unexpected category: '{primary_category}'")
                return self._get_fallback_result("Invalid category returned")
                
        except json.JSONDecodeError as e:
            print(f"⚠️ Failed to parse Gemini JSON response: {e}")
            print(f"Raw response: {response_text}")
            return self._get_fallback_result("JSON parse error")
    
    def _get_fallback_result(self, error_message: str) -> Dict[str, Any]:
        """Get fallback result when Gemini fails"""
        return {
            'category': 'Other/Uncategorized Errors',
            'sub_category': 'LLM Processing Error',
            'confidence': 0,
            'rationale': f'Failed to process with Gemini: {error_message[:100]}'
        }

class LLMService:
    """Main LLM service that manages different providers"""
    
    def __init__(self):
        self.provider = self._get_provider()
        self.rate_limit_delay = 0.1  # Small delay to avoid rate limiting
        self._init_hardcoded_rules()
    
    def _init_hardcoded_rules(self):
        """Initialize hardcoded categorization rules for fast processing"""
        self.hardcoded_rules = {
            # Timeout Errors
            'timeout': ['timeout', 'timed out', 'deadline exceeded', '504', 'gateway timeout', 
                       'request timeout', 'connection timeout', 'operation timeout', 'time out',
                       'timeout error', 'timeout exception', 'read timeout', 'write timeout'],
            
            # Network/Connection Errors
            'network': ['connection failed', 'connection refused', 'connection aborted', 'network unreachable', 
                       'econnreset', 'remote end closed', 'connection error', 'socket error', 'connection lost',
                       'network error', 'connection timeout', 'connection reset', 'connection dropped',
                       'network unreachable', 'host unreachable', 'connection refused', 'connection aborted',
                       'socket timeout', 'connection pool', 'connection limit', 'too many connections'],
            
            # Authentication/Authorization Errors
            'auth': ['unauthorized', 'permission denied', 'authentication failed', '403', 'forbidden', 
                    'invalid api key', 'access denied', 'unauthorized access', 'auth failed', 'login failed',
                    'invalid credentials', 'authentication error', 'authorization failed', 'access forbidden',
                    'invalid token', 'token expired', 'session expired', 'login required', 'auth required'],
            
            # Resource Not Found Errors
            'not_found': ['not found', '404', 'no document selected', 'contains no results', 
                         'resource not found', 'file not found', 'document not found', 'page not found',
                         'no such file or directory', 'errno 2', 'no results found', 'empty result',
                         'no data found', 'missing resource', 'resource missing', 'not available',
                         'no matching', 'no records found', 'empty response', 'no document selected from research'],
            
            # Data Validation/Payload Errors
            'validation': ['invalid data payload', 'validation failed', '400', 'bad request', 
                          'missing required field', 'invalid payload', 'malformed', 'invalid input format',
                          'got "doc" but expected', 'got "pdf" but expected', 'invalid format', 'validation error',
                          'invalid data', 'invalid input', 'invalid parameter', 'invalid argument',
                          'required field missing', 'field validation', 'data validation', 'input validation',
                          'invalid request', 'malformed request', 'bad payload', 'invalid json', 'region_id'],
            
            # Internal Server Errors
            'server': ['internal server error', '500', 'server error', '502', 'bad gateway', 
                      'server overloaded', 'service unavailable', '503', 'service unavailable',
                      'server exception', 'server failure', 'internal error', 'server timeout',
                      'server busy', 'server overloaded', 'service error', 'backend error'],
            
            # LLM Service Errors
            'llm': ['litellm', 'serviceunavailableerror', 'contextwindowexceed', 'rate limit', 
                   'token length exceeds', 'model is currently overloaded', 'openai', 'anthropic',
                   'quota exceeded', 'api error', 'context window exceeded', 'llm error', 'ai error',
                   'model error', 'generation error', 'inference error', 'llm service error',
                   'model unavailable', 'model overloaded', 'ai service error', 'generation failed',
                   'inference failed', 'model timeout', 'ai timeout', 'llm timeout', 'contextwindowexceed',
                   'total token length exceeds', 'allowed limit', 'cannot process more than', 'million tokens',
                   'processing files larger than'],
            
            # Query/Parameter Errors
            'query': ['missing filtertype', 'invalid query', 'parameter', 'sort_by', 'invalid filter',
                     'query error', 'invalid parameter', 'missing parameter', 'invalid query parameter',
                     'filter error', 'search error', 'query failed', 'parameter error', 'invalid sort',
                     'invalid filter type', 'query syntax error', 'malformed query', 'invalid search'],
            
            # Application Exception Errors
            'exception': ['typeerror', 'attributeerror', 'keyerror', 'valueerror', 'nullpointerexception',
                         'object has no attribute', 'cannot unpack', 'nonetype', 'traceback',
                         'an error occured', 'error occured', 'runtimeerror', 'exception', 'python error',
                         'code error', 'programming error', 'application error', 'software error',
                         'bug', 'crash', 'fatal error', 'critical error', 'system error', 'object has no len',
                         'has no attribute', 'object of type', 'nonetype', 'feature_flags'],
            
            # Service Configuration Errors
            'config': ['model configuration unavailable', 'failed to fetch model mapping', 
                      'configuration fetch failed', 'invalid setup', 'configuration error',
                      'setup error', 'config error', 'initialization error', 'startup error',
                      'service configuration', 'config missing', 'invalid configuration',
                      'configuration failed', 'setup failed', 'initialization failed'],
            
            # Data Format Errors
            'format': ['json parse error', 'xml parse error', 'invalid format', 'data structure mismatch',
                      'unexpected token', 'parse error', 'format error', 'parsing error',
                      'invalid json', 'malformed json', 'json error', 'xml error', 'format mismatch',
                      'data format error', 'structure error', 'schema error', 'format validation'],
            
            # Streaming Errors
            'streaming': ['error raised while streaming', 'stream interrupted', 'streaming failed',
                         'stream error', 'streaming error', 'stream timeout', 'stream closed',
                         'streaming timeout', 'stream broken', 'stream failure', 'streaming timeout',
                         'stream interrupted', 'streaming interrupted', 'stream error', 'stream failed',
                         'raised while streaming'],
            
            # Request/Response Logging Errors
            'logging': ['requestid', 'session_id', 'query_id', '{"requestid":', '{"session_id":',
                       'logging data', 'request metadata', 'response data', 'session data',
                       'request log', 'response log', 'audit log', 'access log', 'debug log'],
            
            # Feature Configuration Errors
            'feature': ['feature flag error', 'configuration unavailable', 'feature not enabled',
                       'feature disabled', 'feature flag', 'feature error', 'feature unavailable',
                       'feature not available', 'feature configuration', 'feature setup',
                       'feature initialization', 'feature failed', 'feature timeout']
        }
    
    def _get_provider(self) -> LLMProvider:
        """Get the appropriate LLM provider based on environment configuration"""
        # Check for Azure OpenAI configuration first
        if all([
            os.getenv('AZURE_OPENAI_API_KEY'),
            os.getenv('AZURE_OPENAI_ENDPOINT'),
            os.getenv('AZURE_OPENAI_DEPLOYMENT')
        ]):
            print("🤖 Using Azure OpenAI for error categorization")
            return AzureOpenAIProvider()
        
        # Fallback to Gemini
        elif os.getenv('GEMINI_API_KEY'):
            print("🤖 Using Google Gemini for error categorization")
            return GeminiProvider()
        
        else:
            raise ValueError("No LLM provider configured. Please set up Azure OpenAI or Gemini in .env file")
    
    def _categorize_with_hardcoded_rules(self, error_message: str) -> Optional[str]:
        """Fast hardcoded categorization using keyword matching"""
        error_lower = error_message.lower()
        
        # Check each category
        for category, keywords in self.hardcoded_rules.items():
            for keyword in keywords:
                if keyword in error_lower:
                    # Map internal category names to display names
                    category_map = {
                        'timeout': 'Timeout Errors',
                        'network': 'Network/Connection Errors',
                        'auth': 'Authentication/Authorization Errors',
                        'not_found': 'Resource Not Found Errors',
                        'validation': 'Data Validation/Payload Errors',
                        'server': 'Internal Server Errors',
                        'llm': 'LLM Service Errors',
                        'query': 'Query/Parameter Errors',
                        'exception': 'Application Exception Errors',
                        'config': 'Service Configuration Errors',
                        'format': 'Data Format Errors',
                        'streaming': 'Streaming Errors',
                        'logging': 'Request/Response Logging Errors',
                        'feature': 'Feature Configuration Errors'
                    }
                    return category_map.get(category, 'Other/Uncategorized Errors')
        
        return None  # No hardcoded rule matched
    
    def categorize_error(self, error_message: str) -> str:
        """Categorize an error message and return the primary category"""
        try:
            # Try hardcoded rules first
            hardcoded_category = self._categorize_with_hardcoded_rules(error_message)
            if hardcoded_category:
                return hardcoded_category
            
            # Fall back to LLM if hardcoded rules didn't match
            result = self.provider.categorize_error(error_message)
            category = result.get('category', 'Other/Uncategorized Errors')
            
            # Log the detailed analysis for debugging
            print(f"🔍 LLM Error Analysis: {result.get('sub_category', 'N/A')} (Confidence: {result.get('confidence', 'N/A')}%) - {result.get('rationale', 'N/A')}")
            
            return category
            
        except Exception as e:
            print(f"⚠️ LLM categorization failed: {e}")
            return 'Other/Uncategorized Errors'
    
    def categorize_errors_batch(self, error_messages: list, delay_between_calls: float = 0.1) -> Dict[str, int]:
        """Categorize multiple error messages with hardcoded rules first, then LLM fallback"""
        categories = {}
        unique_errors = list(set(error_messages))  # Remove duplicates for efficiency
        
        print(f"🚀 Categorizing {len(unique_errors)} unique error messages...")
        
        hardcoded_count = 0
        llm_count = 0
        
        for i, error_msg in enumerate(unique_errors):
            try:
                # Try hardcoded rules first
                hardcoded_category = self._categorize_with_hardcoded_rules(error_msg)
                if hardcoded_category:
                    category = hardcoded_category
                    hardcoded_count += 1
                else:
                    # Fall back to LLM for unmatched errors
                    result = self.provider.categorize_error(error_msg)
                    category = result.get('category', 'Other/Uncategorized Errors')
                    llm_count += 1
                    
                    # Log LLM analysis for debugging
                    print(f"🔍 LLM Analysis: {result.get('sub_category', 'N/A')} (Confidence: {result.get('confidence', 'N/A')}%)")
                
                # Count occurrences
                if category in categories:
                    categories[category] += 1
                else:
                    categories[category] = 1
                
                # Add delay to avoid rate limiting (only for LLM calls)
                if llm_count > 0 and llm_count % 10 == 0:
                    time.sleep(delay_between_calls)
                    
            except Exception as e:
                print(f"⚠️ Error categorizing message {i+1}: {e}")
                if 'Other/Uncategorized Errors' in categories:
                    categories['Other/Uncategorized Errors'] += 1
                else:
                    categories['Other/Uncategorized Errors'] = 1
        
        # Print performance summary
        print(f"✅ Categorization complete!")
        print(f"   📊 Hardcoded rules: {hardcoded_count} errors")
        print(f"   🤖 LLM processing: {llm_count} errors")
        if hardcoded_count > 0:
            print(f"   ⚡ Performance gain: {((hardcoded_count / len(unique_errors)) * 100):.1f}% faster")
        
        print(f"📈 Found {len(categories)} error categories.")
        return categories

# Global instance for easy import
llm_service = LLMService()
