"""
WSGI config for graph_api project.

It exposes the WSGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/4.1/howto/deployment/wsgi/
"""

import os

from django.core.wsgi import get_wsgi_application

# Start debugpy inside the WSGI process when DEBUG is enabled so worker
# processes are instrumented and breakpoints are hit.
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'graph_main_app.settings')

if os.environ.get('DEBUG', '').lower() in ('1', 'true'):
    try:
        import debugpy
        debug_port = int(os.environ.get('DEBUGPY_PORT', '5678'))
        debugpy.listen(('0.0.0.0', debug_port))
        if os.environ.get('DEBUG_WAIT_FOR_CLIENT', '').lower() in ('1', 'true'):
            print(f"debugpy listening on 0.0.0.0:{debug_port} (waiting for client)")
            debugpy.wait_for_client()
        else:
            print(f"debugpy listening on 0.0.0.0:{debug_port}")
    except RuntimeError as e:
        # debugpy.listen() already called in this process — ignore
        print('debugpy listen already called:', e)
    except Exception as e:
        # Keep application startup working even if debugpy fails
        print('debugpy failed to start:', e)

application = get_wsgi_application()



