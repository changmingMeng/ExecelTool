# from distutils.core import setup
# import py2exe
#
# includes = ["encodings", "encodings.*"]
# data_files = ['CameraDll.dll']
# options = {"py2exe":
#                {"compressed": 1,
#                 "optimize": 2,
#                 "bundle_files": 1,
#                 "includes": includes
#
#                 }
#            }
# setup(
#    console=['GUI.py'],
#     name='ExecelTool',
#     version='1.0',
#     packages=[''],
#     url='http://www.chinaunicom.com',
#     license='',
#     author='mcm',
#     author_email='471351371@qq.com',
#     description='a simple tool for execl'
# )


from cx_Freeze import setup, Executable
