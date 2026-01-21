import threading
from typing import Set, Any
from .logger import logger

class ProcessRegistry:
    """
    Global registry for tracking active COM process instances (Excel, Word, PowerPoint)
    to ensure they are closed gracefully on application exit or interrupt.
    """
    _instances: list[Any] = []
    _lock = threading.Lock()

    @classmethod
    def register(cls, instance: Any) -> None:
        """Register a COM instance."""
        with cls._lock:
            if instance not in cls._instances:
                cls._instances.append(instance)
                logger.debug(f"Registered process instance: {instance}")

    @classmethod
    def unregister(cls, instance: Any) -> None:
        """Unregister a COM instance."""
        with cls._lock:
            if instance in cls._instances:
                cls._instances.remove(instance)
                logger.debug(f"Unregistered process instance: {instance}")

    @classmethod
    def kill_all(cls) -> None:
        """
        Force close all registered instances. 
        Safe to call from signal handlers or atexit.
        """
        with cls._lock:
            if not cls._instances:
                return
                
            logger.info(f"Cleaning up {len(cls._instances)} active office processes...")
            
            for instance in list(cls._instances):
                try:
                    # Generic COM Quit() method
                    try:
                        instance.Quit()
                    except AttributeError:
                        # Some objects might accept Close() instead, but typically Application objects use Quit()
                        instance.Close()
                        
                    logger.debug(f"Closed process instance: {instance}")
                except Exception as e:
                    logger.warning(f"Failed to close process instance during cleanup: {e}")
            
            cls._instances.clear()
