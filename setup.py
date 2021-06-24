from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
options = {"build_exe": {"packages": ["os"],
                         "include_files": ['settings.txt', 'settings.bak',
                                           'areas.txt', 'assets/'],
                         "optimize": 1},
           "build": {"build_exe": "build/dbmerger_1.0"},
           "install_exe": {"force": True},
           }

target = Executable(
    script="main.py",
    target_name='Merge_Compare',
    base='Console'
)

setup(
    name="Database_Merge_Compare",
    version="1.0",
    description="Merges databases by comparing primary and customer excel sheets",
    author="George M. Marino",
    options=options,
    executables=[target]
)
