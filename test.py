from win32com.client.gencache import EnsureDispatch
import sys
xl = EnsureDispatch("Word.Application")
print(sys.modules[xl.__module__].__file__)