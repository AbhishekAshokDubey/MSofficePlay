import win32com.client
onObj = win32com.client.gencache.EnsureDispatch('OneNote.Application.12')
#result = onObj.GetHierarchy("",win32com.client.constants.hsNotebooks)
result = onObj.GetHierarchy("",win32com.client.constants.hsPages)
print(result)
