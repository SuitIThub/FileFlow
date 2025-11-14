"""Rule classes for file renaming patterns"""
from abc import ABC, abstractmethod
from typing import List, Dict, Any


class Rule(ABC):
    """Abstract base class for renaming rules"""
    
    def __init__(self, tag_name: str):
        self.tag_name = tag_name
    
    @abstractmethod
    def get_value(self, file_index: int, batch_count: int) -> str:
        """Get the replacement value for this rule"""
        pass
    
    @abstractmethod
    def reset(self):
        """Reset the rule state for a new batch"""
        pass
    
    @abstractmethod
    def to_dict(self) -> Dict[str, Any]:
        """Serialize rule to dictionary"""
        pass
    
    @classmethod
    @abstractmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Rule':
        """Deserialize rule from dictionary"""
        pass


class CounterRule(Rule):
    """Rule that counts up with each file in the batch"""
    
    def __init__(self, tag_name: str, start_value: int = 0, increment: int = 1, step: int = 1, max_value: int = None):
        super().__init__(tag_name)
        self.start_value = start_value
        self.increment = increment
        self.step = step  # How many operations before incrementing
        self.max_value = max_value  # Maximum value before wrapping to start_value
        self.current_value = start_value
        self.operation_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        value = self.current_value
        self.operation_count += 1
        
        # Only increment when we've reached the step threshold
        if self.operation_count % self.step == 0:
            self.current_value += self.increment
            
            # Handle max value wrapping
            if self.max_value is not None and self.current_value > self.max_value:
                self.current_value = self.start_value
        
        return str(value)
    
    def reset(self):
        self.current_value = self.start_value
        self.operation_count = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'counter',
            'tag_name': self.tag_name,
            'start_value': self.start_value,
            'increment': self.increment,
            'step': self.step,
            'max_value': self.max_value
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CounterRule':
        return cls(data['tag_name'], data['start_value'], data['increment'], data.get('step', 1), data.get('max_value'))


class ListRule(Rule):
    """Rule that iterates through a list of values"""
    
    def __init__(self, tag_name: str, values: List[str], step: int = 1):
        super().__init__(tag_name)
        self.values = values
        self.step = step  # How many operations before advancing to next value
        self.current_index = 0
        self.operation_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        if not self.values:
            return ""
        
        value = self.values[self.current_index % len(self.values)]
        self.operation_count += 1
        
        # Only advance to next value when we've reached the step threshold
        if self.operation_count % self.step == 0:
            self.current_index += 1
        
        return value
    
    def reset(self):
        self.current_index = 0
        self.operation_count = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'list',
            'tag_name': self.tag_name,
            'values': self.values,
            'step': self.step
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ListRule':
        return cls(data['tag_name'], data['values'], data.get('step', 1))


class BatchRule(Rule):
    """Rule that counts up with each batch"""
    
    def __init__(self, tag_name: str, start_value: int = 0, increment: int = 1, step: int = 1, max_value: int = None):
        super().__init__(tag_name)
        self.start_value = start_value
        self.increment = increment
        self.step = step  # How many batches before incrementing
        self.max_value = max_value  # Maximum value before wrapping to start_value
        self.current_value = start_value
        self.batch_count = 0
    
    def get_value(self, file_index: int, batch_count: int) -> str:
        return str(self.current_value)
    
    def reset(self):
        pass  # Batch counter doesn't reset per batch
    
    def increment_batch(self):
        self.batch_count += 1
        
        # Only increment when we've reached the step threshold
        if self.batch_count % self.step == 0:
            self.current_value += self.increment
            
            # Handle max value wrapping
            if self.max_value is not None and self.current_value > self.max_value:
                self.current_value = self.start_value
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'type': 'batch',
            'tag_name': self.tag_name,
            'start_value': self.start_value,
            'increment': self.increment,
            'step': self.step,
            'max_value': self.max_value,
            'current_value': self.current_value,
            'batch_count': self.batch_count
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BatchRule':
        rule = cls(data['tag_name'], data['start_value'], data['increment'], data.get('step', 1), data.get('max_value'))
        rule.current_value = data.get('current_value', data['start_value'])
        rule.batch_count = data.get('batch_count', 0)
        return rule

