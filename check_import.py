
import sys
import os

# Add current dir to path
sys.path.append(os.getcwd())

try:
    import chemo_filters as cf
    print(f"Successfully imported chemo_filters")
    if hasattr(cf, 'run_comprehensive_screening'):
        print("Function 'run_comprehensive_screening' FOUND")
    else:
        print("Function 'run_comprehensive_screening' NOT FOUND")
        print(f"Available attributes: {dir(cf)}")
except Exception as e:
    print(f"FAILED to import: {e}")
    import traceback
    traceback.print_exc()
