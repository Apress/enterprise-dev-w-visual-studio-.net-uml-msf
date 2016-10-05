function OnFinish(selProj, selObj)
{
    var oldSuppressUIValue = true;
    try
    {
        oldSuppressUIValue = dte.SuppressUI;
        var bSilent = wizard.FindSymbol("SILENT_WIZARD");
        dte.SuppressUI = bSilent;

        var strProjectName = wizard.FindSymbol("PROJECT_NAME");
        var strProjectPath = wizard.FindSymbol("PROJECT_PATH");
        var strTemplatePath = wizard.FindSymbol("TEMPLATES_PATH");
	var strTemplateFile = strTemplatePath + "\\..\\..\\..\\..\\projects\\visual basic building blocks\\" + wizard.FindSymbol("EFT_PROTOTYPE");

        var project = CreateVSProject(strProjectName, ".vbproj", strProjectPath, strTemplateFile);
        if( project )
        {
            strProjectName = project.Name;  //In case it got changed

            var item;

            strTemplateFile = strTemplatePath + "\\Global.asax"; 
            item = AddFileToVSProject("Global.asax", project, project.ProjectItems, strTemplateFile);

            strTemplateFile = strTemplatePath + "\\Web.config"; 
            item = AddFileToVSProject("Web.config", project, project.ProjectItems, strTemplateFile);

            var strRawGuid = wizard.CreateGuid();
            wizard.AddSymbol("GUID_ASSEMBLY", wizard.FormatGuid(strRawGuid, 0));

            strTemplateFile = strTemplatePath + "\\AssemblyInfo.vb"; 
            item = AddFileToVSProject("AssemblyInfo.vb", project, project.ProjectItems, strTemplateFile);

            if( item )
            {
                item.Properties("SubType").Value = "Code";
            }

            strTemplateFile = strTemplatePath + "\\WebService.asmx"; 
            item = AddFileToVSProject("Service1.asmx", project, project.ProjectItems, strTemplateFile);
        
            var configs = new Enumerator(project.ConfigurationManager);
            for(;!configs.atEnd();configs.moveNext())
            {
                configs.item().Properties("StartAction").Value = 0;
                configs.item().Properties("StartPage").Value = "Service1.asmx";
            }

            project.Save();
            CollapseProjectNode( project );
/*            dte.Solution.SolutionBuild.BuildProject( dte.Solution.SolutionBuild.ActiveConfiguration.Name, project.UniqueName ); */
        }
        
        return 0;
    }
    catch(e)
    {   
        switch(e.number)
        {
        case -2147024816 /* FILE_ALREADY_EXISTS */ :
            return -2147213313;

        default:
            SetErrorInfo(e);
            return e.number;
        }
    }
    finally
    {
        dte.SuppressUI = oldSuppressUIValue;
    }
}

function CollapseProjectNode( oProj )
{
    try
    {
	if( !IsSubProject( oProj ) )
        {
            return;
        }
	var UIItemX = GetUIItem( oProj, "" );
	UIItem = UIItemX.UIHierarchyItems;
	UIItem.Expanded = false;
    }
    catch(e)
    {}
}