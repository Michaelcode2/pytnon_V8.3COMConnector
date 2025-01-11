import win32com.client

def list_com_objects():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, '', 0, winreg.KEY_READ) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    class_name = winreg.EnumKey(key, i)
                    if 'V83' in class_name or 'V82' in class_name or '1C' in class_name:
                        print(class_name)
                except:
                    continue
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    list_com_objects() 