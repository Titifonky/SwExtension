using LogDebugging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace SwExtension
{
    public enum ParametreInfo_e
    {
        cTitre = 1,
        cIntitule = 2
    }

    public class Parametre
    {
        private Object _Defaut = null;
        private Object _Valeur = null;
        private Type _Type = null;
        private String _Intitule = "";
        private String _Tip = "";
        private String _Nom = "";

        public T Defaut<T>() { return (T)Convert.ChangeType(_Defaut, typeof(T)); }

        public T GetValeur<T>()
        {
            if (typeof(T).IsEnum)
                return (T)Enum.ToObject(typeof(T), _Valeur);

            return (T)Convert.ChangeType(_Valeur, typeof(T));
        }

        public void SetValeur<T>(T value)
        {
            _Valeur = value;
        }

        public void SetValeur(String value)
        {
            try
            {
                if (_Type.IsEnum)
                    _Valeur = Enum.Parse(_Type, value);
                else
                    _Valeur = Convert.ChangeType(value, _Type);
            }
            catch { _Valeur = _Defaut; }
        }

        public Type Type { get { return _Type; } }
        public String Intitule { get { return _Intitule; } set { _Intitule = value; } }
        public String Tip { get { return _Tip; } set { _Tip = value; } }
        public String Nom { get { return _Nom; } }

        public Parametre(String nom, Type type)
        {
            _Nom = nom;
            _Type = type;
            _Defaut = GetDefaultValue(type);
        }

        public Parametre(String nom, String intitule, Type type)
        {
            _Nom = nom;
            _Intitule = intitule;
            _Type = type;
            _Defaut = GetDefaultValue(type);
        }

        public Parametre(String nom, String intitule, String tip, Type type)
        {
            _Nom = nom;
            _Intitule = intitule;
            _Tip = tip;
            _Type = type;
            _Defaut = GetDefaultValue(type);
        }

        private object GetDefaultValue(Type t)
        {
            if (t.IsValueType && Nullable.GetUnderlyingType(t) == null)
                return Activator.CreateInstance(t);
            else if (t == typeof(String))
                return "";
            else
                return null;
        }
    }

    public class ConfigModule
    {
        private static Boolean _Init = false;
        private static String _XmlPath;
        private static XmlDocument _Doc;
        private static XmlNode _Racine;

        private static readonly String _FichierParametres = "Parametres.xml";
        private static readonly String _TagRacine = "Parametres";
        private static readonly String _TagModule = "Module";
        private static readonly String _TagParametre = "Parametre";

        private static readonly String _TagMNom = "name";
        private static readonly String _TagMIntitule = "intitule";

        private static readonly String _TagPNom = "name";
        private static readonly String _TagPIntitule = "intitule";
        private static readonly String _TagPTip = "tip";
        private static readonly String _TagPType = "type";

        static ConfigModule() { Init(); }

        private static XmlDocument CreerFichier()
        {
            // Création du Xml
            XmlDocument xmlDoc = new XmlDocument();
            // On rajoute la déclaration
            XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xmlDoc.DocumentElement;
            xmlDoc.InsertBefore(xmlDeclaration, root);

            // On ajoute le noeud de base
            XmlNode N = xmlDoc.CreateNode(XmlNodeType.Element, _TagRacine, "");
            xmlDoc.AppendChild(N);

            // On sauve
            xmlDoc.Save(_XmlPath);

            return xmlDoc;
        }

        private static void Init()
        {
            try
            {
                // Chemin du fichier xml
                _XmlPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), _FichierParametres);

                _Doc = new XmlDocument();

                // S'il n'existe pas, on le crée
                if (!File.Exists(_XmlPath))
                    _Doc = CreerFichier();
                else
                    _Doc.Load(_XmlPath);

                // On récupère le noeud de base
                _Racine = _Doc.SelectSingleNode(String.Format("/{0}", _TagRacine));

                // S'il n'existe pas
                if (_Racine == null)
                {
                    // On supprime le fichier
                    File.Delete(_XmlPath);
                    // Et on recommance
                    _Doc = CreerFichier();
                    _Racine = _Doc.SelectSingleNode(String.Format("/{0}", _TagRacine));
                }

                _Init = true;
            }
            catch(Exception e)
            {
                _Init = false;
                Log.Methode(typeof(ConfigModule).ToString(), e);
            }
        }

        private static Boolean EstInit
        {
            get
            {
                // On test
                if (!_Init)
                {
                    // On essaye une deuxième fois
                    Init();

                    if (!_Init)
                    {
                        Log.Methode(typeof(ConfigModule).ToString(), "Le fichier Xml n'est pas chargé");
                        return false;
                    }
                    else
                        return true;
                }

                return true;
            }
        }

        private static XmlNode Module(String nomModule, String intitule, Boolean creerSiExistePas = false)
        {
            if (!EstInit) return null;

            XmlNode NdModule = _Racine.SelectSingleNode(String.Format("{0}[@{1}='{2}']", _TagModule, _TagMNom, nomModule));

            if ((NdModule == null) && creerSiExistePas)
            {
                NdModule = _Doc.CreateNode(XmlNodeType.Element, _TagModule, "");
                XmlAttribute Att = _Doc.CreateAttribute(_TagMNom);
                Att.Value = nomModule;
                NdModule.Attributes.SetNamedItem(Att);

                Att = _Doc.CreateAttribute(_TagMIntitule);
                Att.Value = intitule;
                NdModule.Attributes.SetNamedItem(Att);

                _Racine.AppendChild(NdModule);

                _Doc.Save(_XmlPath);
            }

            if (String.IsNullOrWhiteSpace(GetTag(NdModule, _TagMIntitule)))
            {
                XmlAttribute Att = _Doc.CreateAttribute(_TagMIntitule);
                Att.Value = intitule;
                NdModule.Attributes.SetNamedItem(Att);

                _Doc.Save(_XmlPath);
            }

            return NdModule;
        }

        private static String GetTag(XmlNode n, String nomTag)
        {
            String Val = "";

            if (n.Attributes[nomTag] != null)
                Val = n.Attributes[nomTag].Value;

            return Val;
        }

        private static List<String> ListeTagConfigModule(String Tag)
        {
                List<String> Liste = new List<String>();
                foreach (XmlNode NdModule in _Racine.ChildNodes)
                {
                    if (NdModule.Name != _TagModule) continue;

                    Liste.Add(GetTag(NdModule, Tag));
                }

                return Liste;
        }

        public static List<String> ListeNomConfigModule
        {
            get { return ListeTagConfigModule(_TagMNom); }
        }

        public static List<String> ListeIntituleConfigModule
        {
            get { return ListeTagConfigModule(_TagMIntitule); }
        }

        private String _NomModule = "";
        private String _IntituleModule = "";
        private XmlNode _Module = null;
        private DicParametre _OldDic = new DicParametre();
        private DicParametre _Dic = new DicParametre();

        private class DicParametre
        {
            private Dictionary<String, Parametre> _Dic = null;
            
            public DicParametre()
            {
                _Dic = new Dictionary<string, Parametre>();
            }

            public void Add(Parametre param)
            {
                if(!_Dic.ContainsKey(param.Nom))
                    _Dic.Add(param.Nom, param);
            }

            public void Remove(Parametre param)
            {
                _Dic.Remove(param.Nom);
            }

            public void Remove(String nom)
            {
                _Dic.Remove(nom);
            }

            public Parametre Get(String nom)
            {
                if (_Dic.ContainsKey(nom))
                    return _Dic[nom];

                return null;
            }

            public List<Parametre> Parametres { get { return _Dic.Values.ToList(); } }
        }

        public String Nom { get { return _NomModule; } }

        public String Intitule { get { return _IntituleModule; } }

        public ConfigModule(String nomModule)
        {
            _NomModule = nomModule;
            _IntituleModule = nomModule;

            _Module = Module(nomModule, nomModule, true);

            DicAncienParametre();
        }

        public ConfigModule(String nomModule, String intitule)
        {
            _NomModule = nomModule;
            _IntituleModule = intitule;

            _Module = Module(nomModule, intitule, true);

            DicAncienParametre();
        }

        public void ChargerParametreBrut()
        {
            foreach (XmlNode N in _Module.ChildNodes)
            {
                if (N.Name != _TagParametre) continue;

                Parametre P = new Parametre(GetTag(N, _TagPNom), GetTag(N, _TagPIntitule), GetTag(N, _TagPTip), Type.GetType(GetTag(N, _TagPType)));
                P.SetValeur(N.InnerText);
                _Dic.Add(P);
            }
        }

        private void DicAncienParametre()
        {
            foreach(XmlNode N in _Module.ChildNodes)
            {
                if (N.Name != _TagParametre) continue;

                Parametre P = new Parametre(GetTag(N, _TagPNom), GetTag(N, _TagPIntitule), Type.GetType(GetTag(N, _TagPType)));
                P.SetValeur(N.InnerText);
                _OldDic.Add(P);
            }
        }

        public Parametre AjouterParam<T>(String nom, T valeur, String intitule = "", String tip = "")
        {
            // On recupere le parametre existant
            Parametre P = _OldDic.Get(nom);

            // S'il existe, on met les descriptions à jour
            if((P != null) && (P.Type == typeof(T)))
            {
                P.Intitule = intitule;
                P.Tip = tip;
            }
            // Sinon, on en crée un nouveau
            else
            {
                P = new Parametre(nom, intitule, tip, typeof(T));
                P.SetValeur(valeur);
            }

            _Dic.Add(P);
            _OldDic.Remove(P);
            
            

            return P;
        }

        public Parametre GetParam(String nom)
        {
            Parametre P = _Dic.Get(nom);
            if ((P != null))
                return P;

            this.LogMethode(new Object[] { "Parametre '", nom, "' inconnu" });
            return null;
        }

        public void Sauver()
        {
            _Module.InnerText = "";

            foreach(Parametre P in _Dic.Parametres)
            {
                XmlNode NdParametre = _Doc.CreateNode(XmlNodeType.Element, _TagParametre, "");
                XmlAttribute Att = _Doc.CreateAttribute(_TagPNom);
                Att.Value = P.Nom;
                NdParametre.Attributes.SetNamedItem(Att);

                Att = _Doc.CreateAttribute(_TagPIntitule);
                Att.Value = P.Intitule;
                NdParametre.Attributes.SetNamedItem(Att);

                Att = _Doc.CreateAttribute(_TagPTip);
                Att.Value = P.Tip;
                NdParametre.Attributes.SetNamedItem(Att);

                Att = _Doc.CreateAttribute(_TagPType);
                Att.Value = P.Type.FullName;
                NdParametre.Attributes.SetNamedItem(Att);

                NdParametre.InnerText = P.GetValeur<String>();

                _Module.AppendChild(NdParametre);
            }

            _Doc.Save(_XmlPath);
        }

        public List<Parametre> ListeParametre
        {
            get { return _Dic.Parametres; }
        }
    }
}
