# test_environment.py
import os


def is_running_on_railway():
    """Check if we're running on Railway - more accurate detection"""
    railway_env_vars = [
        "RAILWAY_ENVIRONMENT",
        "RAILWAY_STATIC_URL",
        "RAILWAY_SERVICE_NAME",
    ]
    return any(var in os.environ for var in railway_env_vars)


print("Environment check:")
print(f"RAILWAY_ENVIRONMENT: {os.getenv('RAILWAY_ENVIRONMENT', 'Not set')}")
print(f"RAILWAY_STATIC_URL: {os.getenv('RAILWAY_STATIC_URL', 'Not set')}")
print(f"RAILWAY_SERVICE_NAME: {os.getenv('RAILWAY_SERVICE_NAME', 'Not set')}")
print(f"Is Running on Railway: {is_running_on_railway()}")
