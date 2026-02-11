"""
Configuration loader and exam config manager.

This module provides a clean way to load exam configurations from Python dicts
and create rule instances through dependency injection.
"""

from typing import Dict, List, Any, Optional
from pathlib import Path
from core.rules.base_rule import BaseRule
from core.rules.configurable_rule import ConfigurableRule


class ExamConfig:
    """
    Exam configuration object.
    
    This is what Pipeline expects - a clean interface with rules and output config.
    """
    
    def __init__(self, exam_info: Dict[str, Any], rules: List[BaseRule], output: Dict[str, Any], merge_rules: List[Dict[str, Any]] = None):
        """
        Initialize exam configuration.
        
        Args:
            exam_info: Exam metadata (name, code, etc.)
            rules: List of rule instances
            output: Output configuration (paths, mappings, etc.)
            merge_rules: Optional merge configuration for grouping rules
        """
        self.name = exam_info.get("name", "Unnamed Exam")
        self.code = exam_info.get("code", "unknown")
        self.total_marks = exam_info.get("total_marks", 0)
        self.description = exam_info.get("description", "")
        
        self.rules = rules
        self.output = output
        self.merge_rules = merge_rules or []
    
    def __repr__(self):
        return f"ExamConfig(name='{self.name}', code='{self.code}', rules={len(self.rules)})"


class ConfigLoader:
    """
    Factory for creating ExamConfig from various sources.
    
    This is the central point for loading configurations - supports loading from
    Python dicts, and can be extended to support YAML/JSON/Database in the future.
    """
    
    @staticmethod
    def load_from_dict(config_dict: Dict[str, Any]) -> ExamConfig:
        """
        Load exam configuration from a Python dictionary.
        
        Args:
            config_dict: Configuration dictionary with structure:
                {
                    "exam": {"name": "...", "code": "...", ...},
                    "rules": [
                        {
                            "name": "...",
                            "mark": 1,
                            "checks": [("check_type", {...})],
                            "scoring": "all_or_nothing"
                        }
                    ],
                    "output": {...}
                }
        
        Returns:
            ExamConfig instance ready to use in pipeline
        
        Example:
            >>> config = ConfigLoader.load_from_dict(TEST2_CONFIG)
            >>> run_batch(config, writer)
        """
        # Validate required keys
        if "exam" not in config_dict:
            raise ValueError("Config must have 'exam' key")
        if "rules" not in config_dict:
            raise ValueError("Config must have 'rules' key")
        if "output" not in config_dict:
            raise ValueError("Config must have 'output' key")
        
        # Build rules from configuration
        rules = ConfigLoader._build_rules(config_dict["rules"])
        
        # Extract merge configuration (optional)
        merge_rules = config_dict.get("merge_rules", [])
        
        # Create and return ExamConfig
        return ExamConfig(
            exam_info=config_dict["exam"],
            rules=rules,
            output=config_dict["output"],
            merge_rules=merge_rules
        )
    
    @staticmethod
    def load_from_module(module_name: str, config_var_name: str = None) -> ExamConfig:
        """
        Load configuration from a Python module.
        
        Args:
            module_name: Module path like "config.test2_config"
            config_var_name: Variable name in module (auto-detected if None)
        
        Returns:
            ExamConfig instance
        
        Example:
            >>> config = ConfigLoader.load_from_module("config.test2_config")
        """
        from importlib import import_module
        
        module = import_module(module_name)
        
        # Auto-detect config variable name
        if config_var_name is None:
            # Try common patterns: TEST2_CONFIG, test2_config, CONFIG
            for attr in dir(module):
                if attr.endswith("_CONFIG") or attr == "CONFIG":
                    config_var_name = attr
                    break
            
            if config_var_name is None:
                raise ValueError(f"Cannot find config variable in module {module_name}")
        
        config_dict = getattr(module, config_var_name)
        return ConfigLoader.load_from_dict(config_dict)
    
    @staticmethod
    def _build_rules(rules_config: List[Any]) -> List[BaseRule]:
        """
        Build rule instances from configuration list.
        
        Args:
            rules_config: List of rule configurations (dicts or rule instances)
        
        Returns:
            List of rule instances
        """
        rules = []
        
        for i, rule_config in enumerate(rules_config):
            try:
                # Support backward compatibility: if already a rule instance, use it directly
                if isinstance(rule_config, BaseRule):
                    rules.append(rule_config)
                # Otherwise, treat as dict configuration
                elif isinstance(rule_config, dict):
                    rule = ConfigLoader._create_rule(rule_config)
                    rules.append(rule)
                else:
                    raise TypeError(f"Rule must be a BaseRule instance or dict, got {type(rule_config)}")
            except Exception as e:
                raise ValueError(f"Failed to create rule {i+1}: {e}") from e
        
        return rules
    
    @staticmethod
    def _create_rule(rule_config: Dict[str, Any]) -> BaseRule:
        """
        Create a single rule instance from configuration.
        
        Args:
            rule_config: Rule configuration dict
        
        Returns:
            Rule instance (ConfigurableRule or other BaseRule subclass)
        """
        # Validate required fields
        if "name" not in rule_config:
            raise ValueError("Rule must have 'name' field")
        if "mark" not in rule_config:
            raise ValueError("Rule must have 'mark' field")
        if "checks" not in rule_config:
            raise ValueError("Rule must have 'checks' field")
        
        # Create ConfigurableRule instance
        return ConfigurableRule(
            name=rule_config["name"],
            mark=rule_config["mark"],
            checks=rule_config["checks"],
            scoring=rule_config.get("scoring", "all_or_nothing"),
            min_pass=rule_config.get("min_pass"),
            needs_review_on_error=rule_config.get("needs_review_on_error", False)
        )


def validate_config_dict(config_dict: Dict[str, Any]) -> bool:
    """
    Validate configuration dictionary structure.
    
    Args:
        config_dict: Configuration to validate
    
    Returns:
        True if valid
    
    Raises:
        ValueError: If configuration is invalid
    """
    # Check top-level structure
    required_keys = ["exam", "rules", "output"]
    for key in required_keys:
        if key not in config_dict:
            raise ValueError(f"Missing required key: '{key}'")
    
    # Check exam info
    if not isinstance(config_dict["exam"], dict):
        raise ValueError("'exam' must be a dict")
    
    # Check rules
    if not isinstance(config_dict["rules"], list):
        raise ValueError("'rules' must be a list")
    
    for i, rule in enumerate(config_dict["rules"]):
        if not isinstance(rule, dict):
            raise ValueError(f"Rule {i+1} must be a dict")
        
        required_rule_keys = ["name", "mark", "checks"]
        for key in required_rule_keys:
            if key not in rule:
                raise ValueError(f"Rule {i+1} missing required key: '{key}'")
        
        if not isinstance(rule["checks"], list):
            raise ValueError(f"Rule {i+1} 'checks' must be a list")
    
    # Check output
    if not isinstance(config_dict["output"], dict):
        raise ValueError("'output' must be a dict")
    
    print("✅ Configuration validation passed")
    return True
