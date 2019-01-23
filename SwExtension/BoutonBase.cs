using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using System;
using System.Reflection;

namespace SwExtension
{
    public abstract class BoutonBase : Info
    {
        static BoutonBase() { }

        private Boolean _IsInit = true;
        public Boolean IsInit { get { return _IsInit; } protected set { _IsInit = value; } }

        private eTypeDoc _TypeDocContexteModule = 0;
        public eTypeDoc TypeDocContexteModule
        {
            get { return _TypeDocContexteModule; }
            protected set { _TypeDocContexteModule = value; }
        }

        private String _NomModule = "";
        public String NomModule
        {
            get { return _NomModule; }
            protected set { _NomModule = value; }
        }

        private String _TitreModule = "";
        public String TitreModule
        {
            get { return _TitreModule; }
            protected set { _TitreModule = value; WindowLog.Effacer(); }
        }

        private String _DescriptionModule = "";
        public String DescriptionModule
        {
            get { return _DescriptionModule; }
            protected set { _DescriptionModule = value; }
        }

        private String _AideModule = "";
        public String AideModule
        {
            get { return _AideModule; }
            protected set { _AideModule = value; }
        }

        protected ConfigModule _Config;

        protected ModelDoc2 MdlBase = null;

        public BoutonBase()
        {
            MdlBase = App.ModelDoc2;
            TitreModule = this.GetModuleTitre();
            NomModule = this.GetModuleNom();
            DescriptionModule = this.GetModuleDescription();
            AideModule = this.GetModuleAide();
            TypeDocContexteModule = this.GetModuleTypeDocContexte();

            _Config = new ConfigModule(NomModule, TitreModule);
        }

        protected abstract void Command();

        public void Executer()
        {
            App.MacroEnCours = true;

            AfficherTitre(_TitreModule);

            InitTime();

            Command();

            ExecuterEn(true);
            SautDeLigne();

            App.MacroEnCours = false;
        }

        protected void SauverConfig()
        {
            if (_Config.IsRef()) _Config.Sauver();
        }
    }

    public static class staticModuleInfo
    {
        public static eTypeDoc GetModuleTypeDocContexte(this BoutonBase value)
        {
            return GetModuleTypeDocContexte(value.GetType());
        }

        public static eTypeDoc GetModuleTypeDocContexte(this Type value)
        {
            ModuleTypeDocContexte a = value.GetCustomAttribute<ModuleTypeDocContexte>();
            if (a.IsRef())
                return a.Val;
            else
                return 0;
        }

        public static String GetModuleTitre(this BoutonBase value)
        {
            return GetModuleTitre(value.GetType());
        }

        public static String GetModuleTitre(this Type value)
        {
            ModuleTitre a = value.GetCustomAttribute<ModuleTitre>();
            if (a.IsRef())
                return a.Val;
            else
                return "";
        }

        public static String GetModuleNom(this BoutonBase value)
        {
            return GetModuleNom(value.GetType());
        }

        public static String GetModuleNom(this Type value)
        {
            ModuleNom a = value.GetCustomAttribute<ModuleNom>();
            if (a.IsRef())
                return a.Val;
            else
                return "";
        }

        public static String GetModuleDescription(this BoutonBase value)
        {
            return GetModuleDescription(value.GetType());
        }

        public static String GetModuleDescription(this Type value)
        {
            ModuleDescription a = value.GetCustomAttribute<ModuleDescription>();
            if (a.IsRef())
                return a.Val;
            else
                return "";
        }

        public static String GetModuleAide(this BoutonBase value)
        {
            return GetModuleAide(value.GetType());
        }

        public static String GetModuleAide(this Type value)
        {
            ModuleAide a = value.GetCustomAttribute<ModuleAide>();
            if (a.IsRef())
                return a.Val;
            else
                return "";
        }
    }

    public abstract class ModuleAttribute : System.Attribute { }

    public class ModuleTypeDocContexte : ModuleAttribute
    {
        private eTypeDoc _Val;

        public ModuleTypeDocContexte(eTypeDoc val) { _Val = val; }

        public eTypeDoc Val { get { return _Val; } }

        public static implicit operator eTypeDoc(ModuleTypeDocContexte att) { return att.Val; }
    }

    public class ModuleTitre : ModuleAttribute
    {
        private String _Val;

        public ModuleTitre(String val) { _Val = val; }

        public String Val { get { return _Val; } }

        public static implicit operator String(ModuleTitre att) { return att.Val; }
    }

    public class ModuleNom : ModuleAttribute
    {
        private String _Val;

        public ModuleNom(String val) { _Val = val; }

        public String Val { get { return _Val; } }

        public static implicit operator String(ModuleNom att) { return att.Val; }
    }

    public class ModuleDescription : ModuleAttribute
    {
        private String _Val;

        public ModuleDescription(String val) { _Val = val; }

        public String Val { get { return _Val; } }

        public static implicit operator String(ModuleDescription att) { return att.Val; }
    }

    public class ModuleAide : ModuleAttribute
    {
        private String _Val;

        public ModuleAide(String val) { _Val = val; }

        public String Val { get { return _Val; } }

        public static implicit operator String(ModuleAide att) { return att.Val; }
    }
}
