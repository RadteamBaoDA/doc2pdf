import threading
from typing import Set, Any
from .logger import logger

def kill_office_processes() -> int:
    """Deprecated compatibility no-op; unrelated Office sessions are never killed."""
    logger.warning("kill_office_processes() is disabled for process safety")
    return 0


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
        Force close all registered instances with timeout protection.
        Safe to call from signal handlers or atexit.
        """
        QUIT_TIMEOUT = 5  # seconds per instance
        
        with cls._lock:
            if not cls._instances:
                return
                
            logger.info(f"Cleaning up {len(cls._instances)} active office processes...")
            
            for instance in list(cls._instances):
                # Use timeout to prevent hanging on unresponsive COM objects
                def quit_instance(inst):
                    try:
                        # Try to suppress any dialogs before quitting
                        try:
                            inst.DisplayAlerts = False
                        except:
                            pass
                        
                        # Generic COM Quit() method
                        try:
                            inst.Quit()
                        except AttributeError:
                            # Some objects might accept Close() instead
                            inst.Close()
                            
                        logger.debug(f"Closed process instance: {inst}")
                    except Exception as e:
                        logger.warning(f"Failed to close process instance during cleanup: {e}")
                
                # Run quit with timeout
                thread = threading.Thread(target=quit_instance, args=(instance,))
                thread.daemon = True
                thread.start()
                thread.join(QUIT_TIMEOUT)
                
                if thread.is_alive():
                    logger.warning(f"Process cleanup timed out after {QUIT_TIMEOUT}s - process may need manual termination")
            
            cls._instances.clear()
