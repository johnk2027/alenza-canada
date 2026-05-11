"""
Alenza Capital OS - Monitoring Module
Production monitoring and health checks
"""

import time
import threading
from datetime import datetime, timezone
from typing import Dict, List, Callable
from collections import defaultdict
import psutil
import logging

logger = logging.getLogger(__name__)

class MetricsCollector:
    """Collects and exposes application metrics"""
    
    def __init__(self):
        self.metrics = defaultdict(list)
        self._lock = threading.Lock()
    
    def record_metric(self, name: str, value: float, tags: dict = None):
        """Record a metric value"""
        with self._lock:
            self.metrics[name].append({
                "value": value,
                "timestamp": datetime.now(timezone.utc),
                "tags": tags or {}
            })
    
    def get_metrics(self, name: str = None, minutes: int = 60) -> Dict:
        """Get recent metrics"""
        cutoff = datetime.now(timezone.utc).timestamp() - (minutes * 60)
        
        with self._lock:
            if name:
                return [
                    m for m in self.metrics.get(name, [])
                    if m["timestamp"].timestamp() > cutoff
                ]
            else:
                return {
                    k: [m for m in v if m["timestamp"].timestamp() > cutoff]
                    for k, v in self.metrics.items()
                }

# Global metrics collector
metrics = MetricsCollector()

class PerformanceMonitor:
    """Monitors application performance"""
    
    @staticmethod
    def time_function(func: Callable) -> Callable:
        """Decorator to time function execution"""
        def wrapper(*args, **kwargs):
            start = time.time()
            try:
                result = func(*args, **kwargs)
                duration = time.time() - start
                
                metrics.record_metric(
                    "function_duration",
                    duration,
                    {"function": func.__name__}
                )
                
                if duration > 2.0:
                    logger.warning(
                        f"Slow function: {func.__name__} took {duration:.2f}s"
                    )
                
                return result
            except Exception as e:
                duration = time.time() - start
                metrics.record_metric(
                    "function_error",
                    1,
                    {"function": func.__name__, "error": str(e)}
                )
                raise
        
        return wrapper

class SystemHealth:
    """System health monitoring"""
    
    @staticmethod
    def get_system_stats() -> Dict:
        """Get system resource usage"""
        return {
            "cpu_percent": psutil.cpu_percent(interval=1),
            "memory_percent": psutil.virtual_memory().percent,
            "disk_percent": psutil.disk_usage('/').percent,
            "process_memory_mb": psutil.Process().memory_info().rss / 1024 / 1024
        }
    
    @staticmethod
    def check_database() -> Dict:
        """Check database health"""
        try:
            from database import db_pool
            with db_pool.get_connection() as conn:
                conn.execute("SELECT 1")
                row_count = conn.execute("SELECT COUNT(*) FROM deals").fetchone()[0]
                
            return {
                "status": "healthy",
                "deal_count": row_count,
                "pool_size": len(db_pool._connections)
            }
        except Exception as e:
            return {
                "status": "unhealthy",
                "error": str(e)
            }
    
    @staticmethod
    def check_external_apis() -> Dict:
        """Check external API health"""
        import requests
        
        apis = {
            "bank_of_canada": "https://www.bankofcanada.ca/valet/",
            "statistics_canada": "https://www150.statcan.gc.ca/",
            "corporations_canada": "https://www.ic.gc.ca/"
        }
        
        results = {}
        for name, url in apis.items():
            try:
                response = requests.get(url, timeout=5)
                results[name] = {
                    "status": "available" if response.status_code == 200 else "degraded",
                    "response_time": response.elapsed.total_seconds()
                }
            except:
                results[name] = {"status": "unavailable"}
        
        return results
    
    @staticmethod
    def get_full_health_report() -> Dict:
        """Generate comprehensive health report"""
        return {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "system": SystemHealth.get_system_stats(),
            "database": SystemHealth.check_database(),
            "apis": SystemHealth.check_external_apis(),
            "metrics_summary": {
                k: len(v) for k, v in metrics.metrics.items()
            }
        }
