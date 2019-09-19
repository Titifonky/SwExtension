using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Macros
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Test7"),
        ModuleNom("Test7")]
    public class Test7 : BoutonBase
    {
        public Test7() { }

        protected override void Command()
        {
            try
            {
                var cm = App.CommandManager;

                try
                {
                    WindowLog.Ecrire("Supprimer Tabs");
                    foreach (eTypeDoc T in Enum.GetValues(typeof(eTypeDoc)))
                    {
                        var tabs = (Object[])cm.CommandTabs((int)Sw.eGetSwTypeDoc(T));

                        cm.GetCommandTabCount((int)Sw.eGetSwTypeDoc(T));

                        WindowLog.Ecrire(" " + T.ToString());

                        if (tabs.IsRef())
                        {
                            for (int i = 0; i < tabs.Length; i++)
                            {
                                var tab = tabs[i] as CommandTab;
                                
                                if (tab.IsRef() && tab.GetCommandTabBoxCount() > 0)
                                {
                                    WindowLog.Ecrire("   " + tab.Name);

                                    var ctbs = (Object[])tab.CommandTabBoxes();
                                    if (ctbs.IsRef())
                                    {
                                        for (int j = 0; j < ctbs.Length; j++)
                                        {
                                            var ctb = ctbs[j] as CommandTabBox;
                                            Object Ids;
                                            var r = ctb.GetCommands(out Ids, out _);
                                            WindowLog.Ecrire("     nb buttons : " + r);
                                            if (r == 0) continue;

                                            ctb.RemoveCommands(Ids);
                                            tab.RemoveCommandTabBox(ctb);
                                        }
                                    }
                                }
                                WindowLog.Ecrire("   Supp tab : " + cm.RemoveCommandTab(tab));

                            }
                        }
                    }

                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }

                try
                {
                    WindowLog.SautDeLigne();
                    WindowLog.Ecrire("Command Group");
                    for (int i = 1; i < 65000; i++)
                    {
                        var cg = cm.GetCommandGroup(i);
                        if (cg.IsRef())
                        {
                            WindowLog.Ecrire("   " + cg.Name + " -> " + i);
                            Log.Message("Supprimer CommandGroup : " + (swRemoveCommandGroupErrors)cm.RemoveCommandGroup2(i, false));
                        }
                    }

                }
                catch (Exception e)
                {
                    this.LogMethode(new Object[] { e });
                }

            }
            catch (Exception e)
            {
                this.LogMethode(new Object[] { e });
            }
        }
    }
}
