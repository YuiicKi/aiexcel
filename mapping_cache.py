import json
import os
from datetime import datetime
from typing import List, Dict, Optional

class MappingCache:
    def __init__(self, cache_file: str = 'header_mappings_cache.json'):
        """Initialize mapping cache"""
        self.cache_file = cache_file
        self.cache = self._load_cache()

    def _load_cache(self) -> Dict:
        """Load cache from file"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Error loading cache: {str(e)}")
            return {}

    def _save_cache(self) -> None:
        """Save cache to file"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving cache: {str(e)}")

    def _generate_cache_key(self, headers1: List[str], headers2: List[str]) -> str:
        """Generate unique cache key for header pairs"""
        # Sort headers to ensure consistent key generation
        sorted_headers1 = sorted(headers1)
        sorted_headers2 = sorted(headers2)
        return f"{','.join(sorted_headers1)}|{','.join(sorted_headers2)}"

    def get_mapping(self, headers1: List[str], headers2: List[str], mapping_type: str) -> Optional[Dict[int, int]]:
        """
        Get mapping from cache
        
        Args:
            headers1: First table headers
            headers2: Second table headers
            mapping_type: Type of mapping (e.g., "1_to_2", "1_to_3")
        """
        cache_key = self._generate_cache_key(headers1, headers2)
        if cache_key in self.cache and mapping_type in self.cache[cache_key]:
            mapping_data = self.cache[cache_key][mapping_type]
            # Convert string keys back to integers
            return {int(k): int(v) for k, v in mapping_data.items()}
        return None

    def save_mapping(self, headers1: List[str], headers2: List[str], 
                    mapping_type: str, mapping: Dict[int, int]) -> None:
        """
        Save mapping to cache
        
        Args:
            headers1: First table headers
            headers2: Second table headers
            mapping_type: Type of mapping (e.g., "1_to_2", "1_to_3")
            mapping: Dictionary containing the mapping relationships
        """
        cache_key = self._generate_cache_key(headers1, headers2)
        if cache_key not in self.cache:
            self.cache[cache_key] = {}
        
        # Convert all keys and values to strings for JSON serialization
        mapping_data = {str(k): str(v) for k, v in mapping.items()}
        
        self.cache[cache_key][mapping_type] = mapping_data
        self.cache[cache_key]['last_updated'] = datetime.now().isoformat()
        
        self._save_cache()

    def clear_cache(self) -> None:
        """Clear all cache"""
        self.cache = {}
        self._save_cache()
