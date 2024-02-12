#Enable single node with pop message
def exit_dialog(title = "Operation Notice", content = "Please enable the script switch first!", taskicon = TaskDialogIcon.TaskDialogIconInformation):
    dialog = TaskDialog(title)
    dialog.MainInstruction = content
    dialog.MainIcon = taskicon
    dialog.Show()
    sys.exit("Operation Aborted!")

#Enforce input as list
def input_to_list(obj):
    result = None
    result = obj if isinstance(obj, list) else result = [obj]
    return UnwrapElement(result)
