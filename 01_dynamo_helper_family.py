#Re-load existing families into project
class FamilyOption(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(
            self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source = FamilySource.Family
        overwriteParameterValues = True
        return True
#Alternative way of initialising transaction to family documents
with Transaction(family_document, "Add Parameters") as t:
    try:
        t.Start()
        for definition in definitions:
            family_manager.AddParameter(definition, BuiltInParameterGroup.PG_DATA, True)
        title.append(family_document.Title)
        t.Commit()
    except:
        t.RollBack()
