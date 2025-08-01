"""
LLM Interface module for Revenue Watchdog
Handles AI-powered deal analysis and revenue leakage detection
"""

import json
import requests
from datetime import datetime
from typing import Dict, List, Any
import pandas as pd

from config import DEFAULT_BASE_URL, DEFAULT_MODEL, API_TIMEOUT, HIGH_DISCOUNT_THRESHOLD, OPPORTUNITY_COST_FACTOR
from utils.helpers import is_date_past


class LLMInterface:
    """Handles LLM API calls for deal analysis"""
    
    def __init__(self, api_key: str = "", base_url: str = DEFAULT_BASE_URL):
        self.api_key = api_key
        self.base_url = base_url
        self.model = DEFAULT_MODEL
        
    def analyze_deals(self, parsed_data: List[Dict]) -> Dict[str, Any]:
        """
        Analyze deals data using LLM for revenue leakage detection.
        
        Sample LLM Prompt Design:
        From the following deals data (as JSON), analyze for:
        - Revenue/margin "leak points": unauthorized discounts, missed uplifts, 
          expired incentives, late renewals, out-of-policy credits.
        - For each leak point, output deal ID, risk type, estimated $ impact, 
          and a short remediation suggestion.
        
        Deals Data: {{parsed_data}}
        """
        
        if not self.api_key:
            return self._mock_analysis(parsed_data)
            
        prompt = self._build_analysis_prompt(parsed_data)
        
        try:
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": self.model,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.1,
                    "max_tokens": 2000
                },
                timeout=API_TIMEOUT
            )
            
            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']
                
                # Try to parse JSON from response
                try:
                    return json.loads(content)
                except json.JSONDecodeError:
                    # Fallback if response isn't pure JSON
                    return self._parse_llm_response(content, parsed_data)
            else:
                print(f"API Error: {response.status_code}")
                return self._mock_analysis(parsed_data)
                
        except Exception as e:
            print(f"LLM API Error: {str(e)}")
            return self._mock_analysis(parsed_data)
    
    def _build_analysis_prompt(self, parsed_data: List[Dict]) -> str:
        """Build the analysis prompt for the LLM"""
        return f"""
        Analyze the following deals data for revenue leakage and margin risks:

        ANALYSIS REQUIREMENTS:
        1. Identify leak points: unauthorized discounts, missed uplifts, expired incentives, late renewals, out-of-policy credits
        2. For each issue found, provide: deal_id, risk_type, estimated_impact, remediation_suggestion
        3. Respond in valid JSON format only

        DEALS DATA:
        {json.dumps(parsed_data, indent=2, default=str)}

        Respond with JSON in this format:
        {{
            "summary": {{"total_leakage": 0, "high_risk_deals": 0, "issues_found": 0}},
            "flagged_deals": [
                {{"deal_id": "123", "risk_type": "unauthorized_discount", "impact": 5000, "suggestion": "Review approval process"}}
            ],
            "recommendations": ["action1", "action2"]
        }}
        """
    
    def _mock_analysis(self, parsed_data: List[Dict]) -> Dict[str, Any]:
        """Fallback analysis when LLM is unavailable"""
        flagged_deals = []
        total_leakage = 0
        
        for deal in parsed_data:
            # Basic rule-based analysis
            deal_id = deal.get('deal_id', 'Unknown')
            deal_size = float(deal.get('deal_size', 0))
            discount = float(deal.get('discount_percent', 0))
            
            # Flag high discounts
            if discount > HIGH_DISCOUNT_THRESHOLD:
                impact = deal_size * (discount - HIGH_DISCOUNT_THRESHOLD) / 100
                flagged_deals.append({
                    'deal_id': deal_id,
                    'risk_type': 'unauthorized_discount',
                    'impact': impact,
                    'suggestion': f'Review {discount}% discount approval for deal {deal_id}'
                })
                total_leakage += impact
            
            # Flag expired deals still in pipeline
            close_date = deal.get('close_date', '')
            if close_date and is_date_past(close_date):
                impact = deal_size * OPPORTUNITY_COST_FACTOR
                flagged_deals.append({
                    'deal_id': deal_id,
                    'risk_type': 'phantom_pipeline',
                    'impact': impact,
                    'suggestion': f'Remove expired deal {deal_id} from pipeline'
                })
                total_leakage += impact
        
        return {
            'summary': {
                'total_leakage': total_leakage,
                'high_risk_deals': len(flagged_deals),
                'issues_found': len(flagged_deals)
            },
            'flagged_deals': flagged_deals,
            'recommendations': [
                'Implement discount approval workflow',
                'Set up automated pipeline hygiene alerts',
                'Review pricing strategy for renewals'
            ]
        }
    
    def _parse_llm_response(self, content: str, parsed_data: List[Dict]) -> Dict[str, Any]:
        """Parse non-JSON LLM responses"""
        # TODO: Implement more sophisticated parsing
        return self._mock_analysis(parsed_data)