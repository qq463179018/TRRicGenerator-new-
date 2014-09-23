using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Ric.FileLib.Attribute;
using Ric.FileLib.Entry;
using Ric.FileLib.Enum;
using Ric.FileLib.Exception;
using Ric.FormatLib;

namespace Ric.FileLib
{
    /// <summary>
    /// [Abstract class]
    /// File representation
    ///     | Nda
    ///       > Qa
    ///       > Ia
    ///     | Idn
    ///     | Fm
    ///     | Tc
    ///     | File
    /// </summary>
    public abstract class AFile : IEnumerable<AEntry>, IDisposable, IComparable
    {
        #region Properties

        #region Customization

        /// <summary>
        /// Type of the entry reprentation class 
        /// that will be used for IEnumerable implementation
        /// </summary>
        protected Type EntryType { get; set; }

        /// <summary>
        /// Format of the file
        /// </summary>
        protected AFormat FileFormat { get; set; }

        /// <summary>
        /// Filemode of the file :
        ///     ReadOnly can only be read but not save or changed
        ///     WriteOnly can only be create or save then cannot be read
        ///     ReadWrite can both
        /// </summary>
        protected FileMode Mode { get; set; }

        /// <summary>
        /// Path of the file
        /// </summary>
        public string Path { get; set; }

        #endregion

        #region Content

        /// <summary>
        /// List of titles in the file
        /// </summary>
        protected List<String> Titles { get; set; }

        /// <summary>
        /// Content of the Nda in a dynamic object
        /// </summary>
        /// <remarks>Can fail at Runtime, use with caution</remarks>
        protected List<dynamic> DynamicContent { get; set; }

        /// <summary>
        /// Content of the Nda (excluding title)
        /// Can be accessed via a foreach loop or
        /// a Linq to object query
        /// </summary>
        /// <remarks>See IEnumerable implementation</remarks>
        /// <example>
        /// foreach(NdaEntry entry in myNda) { /* do stuff */ }
        /// from IdnEntry entry in myIdn 
        /// where entry.Ric.StartsWith("6")
        /// select entry
        /// </example>
        public List<AEntry> Content { get; set; }

        #endregion

        #region Checking

        /// <summary>
        /// Functions that will check titles rules
        /// </summary>
        /// <param name="titles"></param>
        /// <returns></returns>
        public delegate void TitleRule(ref List<string> titles);

        /// <summary>
        /// Functions that will check entry rules
        /// </summary>
        /// <param name="entry"></param>
        /// <returns></returns>
        public delegate void EntryRule(ref AEntry entry);

        /// <summary>
        /// Set of rules deleguates to apply to entries
        /// </summary>
        private EntryRule _entryRules;

        /// <summary>
        /// Set of rules deleguates to appy to titles
        /// </summary>
        private TitleRule _titleRules;

        #endregion

        #endregion

        #region Customization

        /// <summary>
        /// Add a rule to check the titles
        /// </summary>
        /// <param name="rule"></param>
        public void AddTitleRule(TitleRule rule)
        {
            _titleRules += rule;
        }

        /// <summary>
        /// Add a rule to check an entry
        /// </summary>
        /// <param name="rule"></param>
        public void AddEntryRule(EntryRule rule)
        {
            _entryRules += rule;
        }

        /// <summary>
        /// Choose the format of the file
        /// Accepted format : 
        ///     | Format.Horizontal
        ///     | Format.Vertical
        ///     | Format.Raw
        /// </summary>
        /// <param name="format">format to choose</param>
        protected void ChooseFormat(Format format)
        {
            if (Equals(format, Format.Horizontal))
            {
                FileFormat = new HorizontalFormat();
            }
            else if (Equals(format, Format.Vertical))
            {
                FileFormat = new VerticalFormat();
            }
            else
            {
                throw new ArgumentException("Format should be Horizontal or Vertical");
            }
        }

        /// <summary>
        /// Set the filemode of the file
        /// Accepted mode ：
        ///     ｜ReadOnly
        ///     ｜WriteOnly
        ///     ｜ReadWrite
        /// </summary>
        /// <param name="mode"></param>
        protected void ChooseMode(FileMode mode)
        {
            Mode = mode;
        }

        /// <summary>
        /// Function to set a custom type
        /// as entry
        /// Need to be a subclass of AEntry
        /// </summary>
        /// <param name="newType"></param>
        public void SetCustomEntryType(Type newType)
        {
            if (newType == null)
            {
                throw new ArgumentNullException("newType", "Type cannot be null");
            }
            if (!newType.IsSubclassOf(typeof(AEntry)))
            {
                throw new FileLibException("Entry representation class should be subClass of AEntry abstract class");
            }
            EntryType = newType;
        }

        #endregion

        #region Loading

        /// <summary>
        /// Load File from an already existing file
        /// </summary>
        /// <param name="path">Path of the Nda file</param>
        public virtual void Load(string path)
        {
            FileFormat.Load(path);
            LoadContent(EntryType);
        }

        /// <summary>
        /// Load File from an already existing file
        /// Using previously set Path
        /// </summary>
        public virtual void Load()
        {
            if (Path == null)
            {
                throw new FileLibException("You need to set path before loading");
            }
            Load(Path);
        }

        /// <summary>
        /// Load file from template and given parameters
        /// Format object will check if templateOrPath is
        /// either a template object or a path of the file
        /// where the template is written
        /// </summary>
        /// <param name="templateOrPath">template or path where the template is</param>
        /// <param name="props"></param>
        protected void LoadFromTemplateObject(object templateOrPath, IEnumerable<Dictionary<string, string>> props)
        {
            FileFormat.LoadTemplate(templateOrPath);
            foreach (var prop in props)
            {
                FileFormat.AddProp(prop);
            }
            FileFormat.Generate();
            LoadContent(EntryType);
        }

        /// <summary>
        /// Loading Nda from Template in file
        /// Accepted template format are
        /// .xls | .xlsx | .csv | .txt
        /// </summary>
        /// <param name="path"></param>
        /// <param name="props"></param>
        public void LoadFromTemplate(string path, List<Dictionary<string, string>> props)
        {
            LoadFromTemplateObject(path, props);
        }

        /// <summary>
        /// Loading Nda from Horizontal Template in F#
        /// </summary>
        /// <param name="template"></param>
        /// <param name="props"></param>
        public void LoadFromTemplate(HFile template, List<Dictionary<string, string>> props)
        {
            LoadFromTemplateObject(template, props);
        }

        /// <summary>
        /// Loading Nda from Vertical Template in F#
        /// </summary>
        /// <param name="template"></param>
        /// <param name="props"></param>
        public void LoadFromTemplate(VFile template, List<Dictionary<string, string>> props)
        {
            LoadFromTemplateObject(template, props);
        }

        #region Loading Content

        /// <summary>
        /// Get the content from Format object
        /// Do conversion between meaningless list
        /// to desired finance file format
        /// //To be implemented for each child file format
        /// </summary>
        protected virtual void LoadContent(Type entryType)
        {
            if (Mode == FileMode.WriteOnly) return;
            LoadContentToFileDynamic();
            LoadContentToFile(entryType);
        }

        /// <summary>
        /// Loading contentm from Format and convert it
        /// to Nda model
        /// </summary>
        private void LoadContentToFile(Type entryType)
        {
            var fullContent = FileFormat.GetContent();
            var enumerable = fullContent as IList<IEnumerable<string>> ?? fullContent.ToList();
            Titles = enumerable.First().ToList();
            foreach (var singleEntry in enumerable.Skip(1))
            {
                var entry = Activator.CreateInstance(entryType);
                var index = 0;
                foreach (var entryPart in singleEntry)
                {
                    var propertyname =
                        (from prop in entryType.GetProperties()
                            from attr in prop.GetCustomAttributes(typeof (TitleName), false)
                            where ((TitleName) attr).Name == Titles[index]
                            select prop.Name).FirstOrDefault();
                    if (propertyname != null)
                    {
                        var propertyType = entryType.GetProperty(propertyname).PropertyType.Name;
                        object value = GetValue(entryPart, propertyType);
                        entryType.GetProperty(propertyname).SetValue(entry, value, null);
                    }
                    index++;
            }
                Content.Add(entry as AEntry);
            }
        }

        /// <summary>
        /// Get object from value
        /// depending the type
        /// </summary>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private object GetValue(string value, string type)
        {
            try
            {
                if (type.ToLower().Contains("string"))
                {
                    return value;
                }
                if (type.ToLower().Contains("datetime"))
                {
                    return DateTime.Parse(value);
                }
                if (type.ToLower().Contains("int"))
                {
                    return Convert.ToInt32(value);
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        /// <summary>
        /// Loading contentm from Format and convert it
        /// to Nda model using dynamic object to represent it
        /// </summary>
        private void LoadContentToFileDynamic()
        {
            var fullContent = FileFormat.GetContent();
            var enumerable = fullContent as IList<IEnumerable<string>> ?? fullContent.ToList();
            Titles = enumerable.First().ToList();
            foreach (var singleNdaEntry in enumerable.Skip(1))
            {
                dynamic ndaEntry = new DynamicEntry();
                var index = 0;
                foreach (var entryPart in singleNdaEntry)
                {
                    ndaEntry.SetProperty(Titles[index].Replace(" ", ""), entryPart);
                    index++;
                }
                DynamicContent.Add(ndaEntry);
            }
        }

        #endregion

        #endregion

        #region Saving

        /// <summary>
        /// Save the generated file
        /// Accepted file format : 
        ///     | .csv
        ///     | .xls / .xlsx
        ///     | .txt
        /// </summary>
        /// <param name="path">path where the file will be saved</param>
        public virtual void Save(string path)
        {
            if (Mode == FileMode.ReadOnly)
            {
                throw new FileLibException("File is Read Only, cannot save");
            }
            if (Mode != FileMode.WriteOnly)
            {
                FileFormat.SetContent(SaveFileToContent());
            }
            Path = path;
            FileFormat.Save(path);
        }

        /// <summary>
        /// Save file using Path parameter
        /// </summary>
        /// <see>
        ///     <cref>Save(string path)</cref>
        /// </see>
        public virtual void Save()
        {
            if (Path == null)
            {
                throw new FileLibException("You need to set path before saving");
            }
            Save(Path);
        }

        /// <summary>
        /// Conversion from Entry to string seq seq
        /// Contrary function of LoadContentToFile()
        /// </summary>
        /// <returns></returns>
        private IEnumerable<IEnumerable<string>> SaveFileToContent()
        {
            var titles = Titles;
            if (_titleRules != null)
            {
                _titleRules(ref titles);
            }

            var content = new List<List<string>> {titles};

            foreach (var aEntry in Content)
            {
                var newEntry = aEntry;
                if (_entryRules != null)
                {
                    _entryRules(ref newEntry);
                }
                var newContentLine = new List<string>();
                foreach (var propName in Titles.Select(
                    title => (from prop in EntryType.GetProperties()
                        from attr in prop.GetCustomAttributes(typeof (TitleName), false)
                        where ((TitleName) attr).Name == title
                        select prop.Name).FirstOrDefault()))
                {
                    if (propName != null)
                    {
                        var contentField = EntryType.GetProperty(propName).GetValue(newEntry, null);
                        newContentLine.Add(contentField.ToString());
                    }
                    else
                    {
                        newContentLine.Add("");
                    }
                }
                content.Add(newContentLine);
            }
            return content;
        }

        #endregion

        #region Title functions

        /// <summary>
        /// Set titles of the file
        /// Cannot be called if ReadOnly
        /// </summary>
        /// <param name="titles"></param>
        public void SetTitles(List<string> titles)
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot SetTitles");
            Titles = titles;
        }

        /// <summary>
        /// Add a list of title to the file
        /// Cannot be called if ReadOnly
        /// </summary>
        /// <param name="titles"></param>
        public void AddTitles(List<string> titles)
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot AddTitles");
            Titles.AddRange(titles);
        }

        /// <summary>
        /// Add a single title
        /// </summary>
        /// <param name="title"></param>
        /// <param name="index"></param>
        public void AddTitle(string title, int index = -1)
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot AddTitle");
            Titles.Add(title);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="titles"></param>
        public void RemoveTitles(List<string> titles)
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot RemoveTitles");
            foreach (var title in titles)
            {
                RemoveTitle(title);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title"></param>
        public void RemoveTitle(string title)
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot RemoveTitle");
            Titles.Remove(title);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<string> GetTitles()
        {
            return Titles;
        }

        /// <summary>
        /// 
        /// </summary>
        public void ResetTitles()
        {
            if (Mode == FileMode.ReadOnly)
                throw new FileLibException("File is Read-Only cannot ResetTitles");
            Titles.Clear();
        }

        #endregion

        #region IEnumerable implementation

        public IEnumerator<AEntry> GetEnumerator()
        {
            if (Mode == FileMode.WriteOnly)
                throw new FileLibException("File is Write-Only cannot iterate");
            return Content.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            if (Mode == FileMode.WriteOnly)
                throw new FileLibException("File is Write-Only cannot iterate");
            return GetEnumerator();
        }

        #endregion

        #region IDisposable implementation

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region IComparable implementation

        public int CompareTo(object obj)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
