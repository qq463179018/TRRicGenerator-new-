// 
//  ____  _     __  __      _        _ 
// |  _ \| |__ |  \/  | ___| |_ __ _| |
// | | | | '_ \| |\/| |/ _ \ __/ _` | |
// | |_| | |_) | |  | |  __/ || (_| | |
// |____/|_.__/|_|  |_|\___|\__\__,_|_|
//
// Auto-generated from deals on 2014-06-30 11:25:43Z.
// Please visit http://code.google.com/p/dblinq2007/ for more information.
//
namespace Ric.Db.Manager
{
    using System;
    using System.ComponentModel;
    using System.Data;
#if MONO_STRICT
	using System.Data.Linq;
#else   // MONO_STRICT
    using DbLinq.Data.Linq;
    using DbLinq.Vendor;
#endif  // MONO_STRICT
    using System.Data.Linq.Mapping;
    using System.Diagnostics;
    using Ric.Db.Info;


    public partial class DealsContext : DataContext
    {
        #region Extensibility Method Declarations
        partial void OnCreated();
        #endregion

        public DealsContext(string connectionString) :
            base(connectionString)
        {
            this.OnCreated();
        }

        public DealsContext(string connection, MappingSource mappingSource) :
            base(connection, mappingSource)
        {
            this.OnCreated();
        }

        public DealsContext(IDbConnection connection, MappingSource mappingSource) :
            base(connection, mappingSource)
        {
            this.OnCreated();
        }

        public Table<CHNProcessItem> Source
        {
            get
            {
                return this.GetTable<CHNProcessItem>();
            }
        }

        public Table<Statistics> Statistics
        {
            get
            {
                return this.GetTable<Statistics>();
            }
        }

        public Table<Configuration> Configuration
        {
            get
            {
                return this.GetTable<Configuration>();
            }
        }
    }

    #region Start MONO_STRICT
#if MONO_STRICT

	public partial class Deals
	{
		
		public Deals(IDbConnection connection) : 
				base(connection)
		{
			this.OnCreated();
		}
	}
    #region End MONO_STRICT
    #endregion
#else     // MONO_STRICT

    public partial class DealsContext
    {
        public DealsContext(IDbConnection connection) :
            base(connection, new DbLinq.MySql.MySqlVendor())
        {
            this.OnCreated();
        }

        public DealsContext(IDbConnection connection, IVendor sqlDialect) :
            base(connection, sqlDialect)
        {
            this.OnCreated();
        }

        public DealsContext(IDbConnection connection, MappingSource mappingSource, IVendor sqlDialect) :
            base(connection, mappingSource, sqlDialect)
        {
            this.OnCreated();
        }
    }
    #region End Not MONO_STRICT
    #endregion
#endif     // MONO_STRICT
    #endregion

    [Table(Name = "deals.source")]
    public partial class JPNProcessItem : ProcessItem, System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged, ICloneable
    {
        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private DateTime _annouceDate;

        private string _assignedTo = string.Empty;

        private string _comments = string.Empty;

        private string _contentKeyword = string.Empty;

        private string _creator = string.Empty;

        private string _dealType = string.Empty;

        private string _englishName = string.Empty;

        private string _filterKeyword = string.Empty;

        private string _htmlPath = string.Empty;

        private int _id;

        private string _isNew = string.Empty;

        private string _logNumber = string.Empty;

        private string _marketName = string.Empty;

        private string _numberOfDeals = string.Empty;

        private string _scopeType = string.Empty;

        private string _sourceFrom = string.Empty;

        private string _sourceLink = string.Empty;

        private string _status = string.Empty;

        private string _targetNation = string.Empty;

        private string _ticker = string.Empty;

        private string _time = string.Empty;

        private string _title = string.Empty;

        private string _sourcingEngine = string.Empty;

        private DateTime _createDate;

        private DateTime _lastModify;

        private DateTime _finishedDate;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnAnnouceDateChanged();

        partial void OnAnnouceDateChanging(DateTime value);

        partial void OnAssignedToChanged();

        partial void OnAssignedToChanging(string value);

        partial void OnCommentsChanged();

        partial void OnCommentsChanging(string value);

        partial void OnContentKeywordChanged();

        partial void OnContentKeywordChanging(string value);

        partial void OnCreatorChanged();

        partial void OnCreatorChanging(string value);

        partial void OnDealTypeChanged();

        partial void OnDealTypeChanging(string value);

        partial void OnEnglishNameChanged();

        partial void OnEnglishNameChanging(string value);

        partial void OnFilterKeywordChanged();

        partial void OnFilterKeywordChanging(string value);

        partial void OnHtmlPathChanged();

        partial void OnHtmlPathChanging(string value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIsNewChanged();

        partial void OnIsNewChanging(string value);

        partial void OnLogNumberChanged();

        partial void OnLogNumberChanging(string value);

        partial void OnMarketNameChanged();

        partial void OnMarketNameChanging(string value);

        partial void OnNumberOfDealsChanged();

        partial void OnNumberOfDealsChanging(string value);

        partial void OnScopeTypeChanged();

        partial void OnScopeTypeChanging(string value);

        partial void OnSourceFromChanged();

        partial void OnSourceFromChanging(string value);

        partial void OnSourceLinkChanged();

        partial void OnSourceLinkChanging(string value);

        partial void OnStatusChanged();

        partial void OnStatusChanging(string value);

        partial void OnTargetNationChanged();

        partial void OnTargetNationChanging(string value);

        partial void OnTickerChanged();

        partial void OnTickerChanging(string value);

        partial void OnTimeChanged();

        partial void OnTimeChanging(string value);

        partial void OnTitleChanged();

        partial void OnTitleChanging(string value);

        partial void OnSourcingEngineChanged();

        partial void OnSourcingEngineChanging(string value);

        partial void OnCreateDateChanged();

        partial void OnCreateDateChanging(DateTime value);

        partial void OnLastModifyChanged();

        partial void OnLastModifyChanging(DateTime value);

        partial void OnFinishedDateChanged();

        partial void OnFinishedDateChanging(DateTime value);
        #endregion


        public JPNProcessItem()
        {
            this.OnCreated();
        }

        [Column(Storage = "_createDate", Name = "CreateDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime CreateDate
        {
            get
            {
                return _createDate;
            }
            set
            {
                if (((_createDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._createDate = value;
                    this.SendPropertyChanged("CreateDate");
                    this.OnCreateDateChanged();
                }
            }
        }

        [Column(Storage = "_finishedDate", Name = "FinishedDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime FinishedDate
        {
            get
            {
                return _finishedDate;
            }
            set
            {
                if (((_finishedDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._finishedDate = value;
                    this.SendPropertyChanged("FinishedDate");
                    this.OnFinishedDateChanged();
                }
            }
        }

        public string CreateDate_
        {
            get
            {
                return CreateDate.ToString();
            }
        }

        public string LastModify_
        {
            get
            {
                return LastModify.ToString();
            }
        }

        public string AnnouceDate_
        {
            get
            {
                return AnnouceDate.ToString();
            }
        }

        [Column(Storage = "_lastModify", Name = "LastModify", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime LastModify
        {
            get
            {
                return _lastModify;
            }
            set
            {
                if (((_lastModify == value)
                            == false))
                {
                    this.OnLastModifyChanging(value);
                    this.SendPropertyChanging();
                    this._lastModify = value;
                    this.SendPropertyChanged("LastModify");
                    this.OnLastModifyChanged();
                }
            }
        }

        [Column(Storage = "_annouceDate", Name = "AnnouceDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime AnnouceDate
        {
            get
            {
                return this._annouceDate;
            }
            set
            {
                if (((_annouceDate == value)
                            == false))
                {
                    this.OnAnnouceDateChanging(value);
                    this.SendPropertyChanging();
                    this._annouceDate = value;
                    this.SendPropertyChanged("AnnouceDate");
                    this.OnAnnouceDateChanged();
                }
            }
        }

        [Column(Storage = "_assignedTo", Name = "AssignedTo", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string AssignedTo
        {
            get
            {
                return this._assignedTo;
            }
            set
            {
                if (((_assignedTo == value)
                            == false))
                {
                    this.OnAssignedToChanging(value);
                    this.SendPropertyChanging();
                    this._assignedTo = value;
                    this.SendPropertyChanged("AssignedTo");
                    this.OnAssignedToChanged();
                }
            }
        }

        [Column(Storage = "_comments", Name = "Comments", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Comments
        {
            get
            {
                return this._comments;
            }
            set
            {
                if (((_comments == value)
                            == false))
                {
                    this.OnCommentsChanging(value);
                    this.SendPropertyChanging();
                    this._comments = value;
                    this.SendPropertyChanged("Comments");
                    this.OnCommentsChanged();
                }
            }
        }

        [Column(Storage = "_contentKeyword", Name = "ContentKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ContentKeyword
        {
            get
            {
                return this._contentKeyword;
            }
            set
            {
                if (((_contentKeyword == value)
                            == false))
                {
                    this.OnContentKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._contentKeyword = value;
                    this.SendPropertyChanged("ContentKeyword");
                    this.OnContentKeywordChanged();
                }
            }
        }

        [Column(Storage = "_creator", Name = "Creator", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Creator
        {
            get
            {
                return this._creator;
            }
            set
            {
                if (((_creator == value)
                            == false))
                {
                    this.OnCreatorChanging(value);
                    this.SendPropertyChanging();
                    this._creator = value;
                    this.SendPropertyChanged("Creator");
                    this.OnCreatorChanged();
                }
            }
        }

        [Column(Storage = "_dealType", Name = "DealType", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string DealType
        {
            get
            {
                return this._dealType;
            }
            set
            {
                if (((_dealType == value)
                            == false))
                {
                    this.OnDealTypeChanging(value);
                    this.SendPropertyChanging();
                    this._dealType = value;
                    this.SendPropertyChanged("DealType");
                    this.OnDealTypeChanged();
                }
            }
        }

        [Column(Storage = "_englishName", Name = "EnglishName", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string EnglishName
        {
            get
            {
                return this._englishName;
            }
            set
            {
                if (((_englishName == value)
                            == false))
                {
                    this.OnEnglishNameChanging(value);
                    this.SendPropertyChanging();
                    this._englishName = value;
                    this.SendPropertyChanged("EnglishName");
                    this.OnEnglishNameChanged();
                }
            }
        }

        [Column(Storage = "_filterKeyword", Name = "FilterKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string FilterKeyword
        {
            get
            {
                return this._filterKeyword;
            }
            set
            {
                if (((_filterKeyword == value)
                            == false))
                {
                    this.OnFilterKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._filterKeyword = value;
                    this.SendPropertyChanged("FilterKeyword");
                    this.OnFilterKeywordChanged();
                }
            }
        }

        [Column(Storage = "_htmlPath", Name = "HtmlPath", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string HtmlPath
        {
            get
            {
                return this._htmlPath;
            }
            set
            {
                if (((_htmlPath == value)
                            == false))
                {
                    this.OnHtmlPathChanging(value);
                    this.SendPropertyChanging();
                    this._htmlPath = value;
                    this.SendPropertyChanged("HtmlPath");
                    this.OnHtmlPathChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "Id", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_isNew", Name = "IsNew", DbType = "varchar(10)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string IsNew
        {
            get
            {
                return this._isNew;
            }
            set
            {
                if ((_isNew != value))
                {
                    this.OnIsNewChanging(value);
                    this.SendPropertyChanging();
                    this._isNew = value;
                    this.SendPropertyChanged("IsNew");
                    this.OnIsNewChanged();
                }
            }
        }

        [Column(Storage = "_logNumber", Name = "LogNumber", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string LogNumber
        {
            get
            {
                return this._logNumber;
            }
            set
            {
                if (((_logNumber == value)
                            == false))
                {
                    this.OnLogNumberChanging(value);
                    this.SendPropertyChanging();
                    this._logNumber = value;
                    this.SendPropertyChanged("LogNumber");
                    this.OnLogNumberChanged();
                }
            }
        }

        [Column(Storage = "_marketName", Name = "MarketName", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string MarketName
        {
            get
            {
                return this._marketName;
            }
            set
            {
                if (((_marketName == value)
                            == false))
                {
                    this.OnMarketNameChanging(value);
                    this.SendPropertyChanging();
                    this._marketName = value;
                    this.SendPropertyChanged("MarketName");
                    this.OnMarketNameChanged();
                }
            }
        }

        [Column(Storage = "_numberOfDeals", Name = "NumberOfDeals", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string NumberOfDeals
        {
            get
            {
                return this._numberOfDeals;
            }
            set
            {
                if ((_numberOfDeals != value))
                {
                    this.OnNumberOfDealsChanging(value);
                    this.SendPropertyChanging();
                    this._numberOfDeals = value;
                    this.SendPropertyChanged("NumberOfDeals");
                    this.OnNumberOfDealsChanged();
                }
            }
        }

        [Column(Storage = "_scopeType", Name = "ScopeType", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ScopeType
        {
            get
            {
                return this._scopeType;
            }
            set
            {
                if (((_scopeType == value)
                            == false))
                {
                    this.OnScopeTypeChanging(value);
                    this.SendPropertyChanging();
                    this._scopeType = value;
                    this.SendPropertyChanged("ScopeType");
                    this.OnScopeTypeChanged();
                }
            }
        }

        [Column(Storage = "_sourceFrom", Name = "SourceFrom", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceFrom
        {
            get
            {
                return this._sourceFrom;
            }
            set
            {
                if (((_sourceFrom == value)
                            == false))
                {
                    this.OnSourceFromChanging(value);
                    this.SendPropertyChanging();
                    this._sourceFrom = value;
                    this.SendPropertyChanged("SourceFrom");
                    this.OnSourceFromChanged();
                }
            }
        }

        [Column(Storage = "_sourceLink", Name = "SourceLink", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceLink
        {
            get
            {
                return this._sourceLink;
            }
            set
            {
                if (((_sourceLink == value)
                            == false))
                {
                    this.OnSourceLinkChanging(value);
                    this.SendPropertyChanging();
                    this._sourceLink = value;
                    this.SendPropertyChanged("SourceLink");
                    this.OnSourceLinkChanged();
                }
            }
        }

        [Column(Storage = "_status", Name = "Status", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Status
        {
            get
            {
                return this._status;
            }
            set
            {
                if (((_status == value)
                            == false))
                {
                    this.OnStatusChanging(value);
                    this.SendPropertyChanging();
                    this._status = value;
                    this.SendPropertyChanged("Status");
                    this.OnStatusChanged();
                }
            }
        }

        [Column(Storage = "_targetNation", Name = "TargetNation", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string TargetNation
        {
            get
            {
                return this._targetNation;
            }
            set
            {
                if (((_targetNation == value)
                            == false))
                {
                    this.OnTargetNationChanging(value);
                    this.SendPropertyChanging();
                    this._targetNation = value;
                    this.SendPropertyChanged("TargetNation");
                    this.OnTargetNationChanged();
                }
            }
        }

        [Column(Storage = "_ticker", Name = "Ticker", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Ticker
        {
            get
            {
                return this._ticker;
            }
            set
            {
                if (((_ticker == value)
                            == false))
                {
                    this.OnTickerChanging(value);
                    this.SendPropertyChanging();
                    this._ticker = value;
                    this.SendPropertyChanged("Ticker");
                    this.OnTickerChanged();
                }
            }
        }

        [Column(Storage = "_time", Name = "Time", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Time
        {
            get
            {
                return this._time;
            }
            set
            {
                if (((_time == value)
                            == false))
                {
                    this.OnTimeChanging(value);
                    this.SendPropertyChanging();
                    this._time = value;
                    this.SendPropertyChanged("Time");
                    this.OnTimeChanged();
                }
            }
        }

        [Column(Storage = "_title", Name = "Title", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Title
        {
            get
            {
                return this._title;
            }
            set
            {
                if (((_title == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._title = value;
                    this.SendPropertyChanged("Title");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_sourcingEngine", Name = "SourcingEngine", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourcingEngine
        {
            get
            {
                return this._sourcingEngine;
            }
            set
            {
                if (((_sourcingEngine == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._sourcingEngine = value;
                    this.SendPropertyChanged("SourcingEngine");
                    this.OnTitleChanged();
                }
            }
        }

        public string PDF_LocalPath { get; set; }
        public string TXT_LocalPath { get; set; }
        public string MarketCode { get; set; }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        #region ICloneable Members

        public object Clone()
        {
            JPNProcessItem item = new JPNProcessItem();
            item.AssignedTo = this.AssignedTo;
            item.Comments = this.Comments;
            item.Creator = this.Creator;
            item.DealType = this.DealType;
            item.SourceFrom = this.SourceFrom;
            item.EnglishName = this.EnglishName;
            item.ID = this.ID;
            item.IsNew = this.IsNew;
            item.LogNumber = this.LogNumber;
            item.MarketName = this.MarketName;
            item.NumberOfDeals = this.NumberOfDeals;
            item.ScopeType = this.ScopeType;
            item.AnnouceDate = this.AnnouceDate;
            item.Time = this.Time;
            item.SourceLink = this.SourceLink;
            item.Status = this.Status;
            item.TargetNation = this.TargetNation;
            item.Ticker = this.Ticker;
            item.Title = this.Title;
            item.MarketName = this.MarketName;
            item.FilterKeyword = this.FilterKeyword;
            item.ContentKeyword = this.ContentKeyword;
            item.HtmlPath = this.HtmlPath;
            item.CreateDate = this.CreateDate;
            item.LastModify = this.LastModify;

            item.MarketCode = this.MarketCode;
            item.PDF_LocalPath = this.PDF_LocalPath;
            item.TXT_LocalPath = this.TXT_LocalPath;

            item.Content = this.Content;
            item.Url = this.Url;
            item.CaptureDate = this.CaptureDate;
            item.SourcingEngine = this.SourcingEngine;
            item.FinishedDate = this.FinishedDate;
            return item;
        }

        #endregion
    }

    [Table(Name = "deals.source")]
    public partial class CHNProcessItem : ProcessItem, System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged, ICloneable
    {
        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private DateTime _annouceDate;

        private string _assignedTo = string.Empty;

        private string _comments = string.Empty;

        private string _contentKeyword = string.Empty;

        private string _creator = string.Empty;

        private string _dealType = string.Empty;

        private string _englishName = string.Empty;

        private string _filterKeyword = string.Empty;

        private string _htmlPath = string.Empty;

        private int _id;

        private string _isNew = string.Empty;

        private string _logNumber = string.Empty;

        private string _marketName = string.Empty;

        private string _numberOfDeals = string.Empty;

        private string _scopeType = string.Empty;

        private string _sourceFrom = string.Empty;

        private string _sourceLink = string.Empty;

        private string _status = string.Empty;

        private string _targetNation = string.Empty;

        private string _ticker = string.Empty;

        private string _time = string.Empty;

        private string _title = string.Empty;

        private string _sourcingEngine = string.Empty;

        private DateTime _createDate;

        private DateTime _lastModify;

        private DateTime _finishedDate;

        private string _peComments = string.Empty;

        private string _sourceOfArticle = string.Empty;

        private string _lastModifiedBy = string.Empty;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnAnnouceDateChanged();

        partial void OnAnnouceDateChanging(DateTime value);

        partial void OnAssignedToChanged();

        partial void OnAssignedToChanging(string value);

        partial void OnCommentsChanged();

        partial void OnCommentsChanging(string value);

        partial void OnContentKeywordChanged();

        partial void OnContentKeywordChanging(string value);

        partial void OnCreatorChanged();

        partial void OnCreatorChanging(string value);

        partial void OnDealTypeChanged();

        partial void OnDealTypeChanging(string value);

        partial void OnEnglishNameChanged();

        partial void OnEnglishNameChanging(string value);

        partial void OnFilterKeywordChanged();

        partial void OnFilterKeywordChanging(string value);

        partial void OnHtmlPathChanged();

        partial void OnHtmlPathChanging(string value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIsNewChanged();

        partial void OnIsNewChanging(string value);

        partial void OnLogNumberChanged();

        partial void OnLogNumberChanging(string value);

        partial void OnMarketNameChanged();

        partial void OnMarketNameChanging(string value);

        partial void OnNumberOfDealsChanged();

        partial void OnNumberOfDealsChanging(string value);

        partial void OnScopeTypeChanged();

        partial void OnScopeTypeChanging(string value);

        partial void OnSourceFromChanged();

        partial void OnSourceFromChanging(string value);

        partial void OnSourceLinkChanged();

        partial void OnSourceLinkChanging(string value);

        partial void OnStatusChanged();

        partial void OnStatusChanging(string value);

        partial void OnTargetNationChanged();

        partial void OnTargetNationChanging(string value);

        partial void OnTickerChanged();

        partial void OnTickerChanging(string value);

        partial void OnTimeChanged();

        partial void OnTimeChanging(string value);

        partial void OnTitleChanged();

        partial void OnTitleChanging(string value);

        partial void OnSourcingEngineChanged();

        partial void OnSourcingEngineChanging(string value);

        partial void OnCreateDateChanged();

        partial void OnCreateDateChanging(DateTime value);

        partial void OnLastModifyChanged();

        partial void OnLastModifyChanging(DateTime value);

        partial void OnFinishedDateChanged();

        partial void OnFinishedDateChanging(DateTime value);

        partial void OnPeCommentsChanged();

        partial void OnPeCommentsChanging(string value);

        partial void OnSourceOfArticleChanged();

        partial void OnSourceOfArticleChanging(string value);

        partial void OnLastModifiedByChanged();

        partial void OnLastModifiedByChanging(string value);
        #endregion


        public CHNProcessItem()
        {
            this.OnCreated();
        }

        [Column(Storage = "_createDate", Name = "CreateDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime CreateDate
        {
            get
            {
                return _createDate;
            }
            set
            {
                if (((_createDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._createDate = value;
                    this.SendPropertyChanged("CreateDate");
                    this.OnCreateDateChanged();
                }
            }
        }

        [Column(Storage = "_finishedDate", Name = "FinishedDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime FinishedDate
        {
            get
            {
                return _finishedDate;
            }
            set
            {
                if (((_finishedDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._finishedDate = value;
                    this.SendPropertyChanged("FinishedDate");
                    this.OnFinishedDateChanged();
                }
            }
        }

        public string CreateDate_
        {
            get
            {
                return CreateDate.ToString();
            }
        }

        public string LastModify_
        {
            get
            {
                return LastModify.ToString();
            }
        }

        public string AnnouceDate_
        {
            get
            {
                return AnnouceDate.ToString();
            }
        }

        [Column(Storage = "_lastModify", Name = "LastModify", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime LastModify
        {
            get
            {
                return _lastModify;
            }
            set
            {
                if (((_lastModify == value)
                            == false))
                {
                    this.OnLastModifyChanging(value);
                    this.SendPropertyChanging();
                    this._lastModify = value;
                    this.SendPropertyChanged("LastModify");
                    this.OnLastModifyChanged();
                }
            }
        }

        [Column(Storage = "_annouceDate", Name = "AnnouceDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime AnnouceDate
        {
            get
            {
                return this._annouceDate;
            }
            set
            {
                if (((_annouceDate == value)
                            == false))
                {
                    this.OnAnnouceDateChanging(value);
                    this.SendPropertyChanging();
                    this._annouceDate = value;
                    this.SendPropertyChanged("AnnouceDate");
                    this.OnAnnouceDateChanged();
                }
            }
        }

        [Column(Storage = "_assignedTo", Name = "AssignedTo", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string AssignedTo
        {
            get
            {
                return this._assignedTo;
            }
            set
            {
                if (((_assignedTo == value)
                            == false))
                {
                    this.OnAssignedToChanging(value);
                    this.SendPropertyChanging();
                    this._assignedTo = value;
                    this.SendPropertyChanged("AssignedTo");
                    this.OnAssignedToChanged();
                }
            }
        }

        [Column(Storage = "_comments", Name = "Comments", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Comments
        {
            get
            {
                return this._comments;
            }
            set
            {
                if (((_comments == value)
                            == false))
                {
                    this.OnCommentsChanging(value);
                    this.SendPropertyChanging();
                    this._comments = value;
                    this.SendPropertyChanged("Comments");
                    this.OnCommentsChanged();
                }
            }
        }

        [Column(Storage = "_peComments", Name = "PeComments", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string PeComments
        {
            get
            {
                return this._peComments;
            }
            set
            {
                if (((_peComments == value)
                            == false))
                {
                    this.OnPeCommentsChanging(value);
                    this.SendPropertyChanging();
                    this._peComments = value;
                    this.SendPropertyChanged("PeComments");
                    this.OnPeCommentsChanged();
                }
            }
        }

        [Column(Storage = "_contentKeyword", Name = "ContentKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ContentKeyword
        {
            get
            {
                return this._contentKeyword;
            }
            set
            {
                if (((_contentKeyword == value)
                            == false))
                {
                    this.OnContentKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._contentKeyword = value;
                    this.SendPropertyChanged("ContentKeyword");
                    this.OnContentKeywordChanged();
                }
            }
        }

        [Column(Storage = "_creator", Name = "Creator", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Creator
        {
            get
            {
                return this._creator;
            }
            set
            {
                if (((_creator == value)
                            == false))
                {
                    this.OnCreatorChanging(value);
                    this.SendPropertyChanging();
                    this._creator = value;
                    this.SendPropertyChanged("Creator");
                    this.OnCreatorChanged();
                }
            }
        }

        [Column(Storage = "_dealType", Name = "DealType", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string DealType
        {
            get
            {
                return this._dealType;
            }
            set
            {
                if (((_dealType == value)
                            == false))
                {
                    this.OnDealTypeChanging(value);
                    this.SendPropertyChanging();
                    this._dealType = value;
                    this.SendPropertyChanged("DealType");
                    this.OnDealTypeChanged();
                }
            }
        }

        [Column(Storage = "_englishName", Name = "EnglishName", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string EnglishName
        {
            get
            {
                return this._englishName;
            }
            set
            {
                if (((_englishName == value)
                            == false))
                {
                    this.OnEnglishNameChanging(value);
                    this.SendPropertyChanging();
                    this._englishName = value;
                    this.SendPropertyChanged("EnglishName");
                    this.OnEnglishNameChanged();
                }
            }
        }

        [Column(Storage = "_filterKeyword", Name = "FilterKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string FilterKeyword
        {
            get
            {
                return this._filterKeyword;
            }
            set
            {
                if (((_filterKeyword == value)
                            == false))
                {
                    this.OnFilterKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._filterKeyword = value;
                    this.SendPropertyChanged("FilterKeyword");
                    this.OnFilterKeywordChanged();
                }
            }
        }

        [Column(Storage = "_htmlPath", Name = "HtmlPath", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string HtmlPath
        {
            get
            {
                return this._htmlPath;
            }
            set
            {
                if (((_htmlPath == value)
                            == false))
                {
                    this.OnHtmlPathChanging(value);
                    this.SendPropertyChanging();
                    this._htmlPath = value;
                    this.SendPropertyChanged("HtmlPath");
                    this.OnHtmlPathChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "Id", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_isNew", Name = "IsNew", DbType = "varchar(10)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string IsNew
        {
            get
            {
                return this._isNew;
            }
            set
            {
                if ((_isNew != value))
                {
                    this.OnIsNewChanging(value);
                    this.SendPropertyChanging();
                    this._isNew = value;
                    this.SendPropertyChanged("IsNew");
                    this.OnIsNewChanged();
                }
            }
        }

        [Column(Storage = "_logNumber", Name = "LogNumber", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string LogNumber
        {
            get
            {
                return this._logNumber;
            }
            set
            {
                if (((_logNumber == value)
                            == false))
                {
                    this.OnLogNumberChanging(value);
                    this.SendPropertyChanging();
                    this._logNumber = value;
                    this.SendPropertyChanged("LogNumber");
                    this.OnLogNumberChanged();
                }
            }
        }

        [Column(Storage = "_marketName", Name = "MarketName", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string MarketName
        {
            get
            {
                return this._marketName;
            }
            set
            {
                if (((_marketName == value)
                            == false))
                {
                    this.OnMarketNameChanging(value);
                    this.SendPropertyChanging();
                    this._marketName = value;
                    this.SendPropertyChanged("MarketName");
                    this.OnMarketNameChanged();
                }
            }
        }

        [Column(Storage = "_numberOfDeals", Name = "NumberOfDeals", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string NumberOfDeals
        {
            get
            {
                return this._numberOfDeals;
            }
            set
            {
                if ((_numberOfDeals != value))
                {
                    this.OnNumberOfDealsChanging(value);
                    this.SendPropertyChanging();
                    this._numberOfDeals = value;
                    this.SendPropertyChanged("NumberOfDeals");
                    this.OnNumberOfDealsChanged();
                }
            }
        }

        [Column(Storage = "_scopeType", Name = "ScopeType", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ScopeType
        {
            get
            {
                return this._scopeType;
            }
            set
            {
                if (((_scopeType == value)
                            == false))
                {
                    this.OnScopeTypeChanging(value);
                    this.SendPropertyChanging();
                    this._scopeType = value;
                    this.SendPropertyChanged("ScopeType");
                    this.OnScopeTypeChanged();
                }
            }
        }

        [Column(Storage = "_sourceFrom", Name = "SourceFrom", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceFrom
        {
            get
            {
                return this._sourceFrom;
            }
            set
            {
                if (((_sourceFrom == value)
                            == false))
                {
                    this.OnSourceFromChanging(value);
                    this.SendPropertyChanging();
                    this._sourceFrom = value;
                    this.SendPropertyChanged("SourceFrom");
                    this.OnSourceFromChanged();
                }
            }
        }

        [Column(Storage = "_sourceLink", Name = "SourceLink", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceLink
        {
            get
            {
                return this._sourceLink;
            }
            set
            {
                if (((_sourceLink == value)
                            == false))
                {
                    this.OnSourceLinkChanging(value);
                    this.SendPropertyChanging();
                    this._sourceLink = value;
                    this.SendPropertyChanged("SourceLink");
                    this.OnSourceLinkChanged();
                }
            }
        }

        [Column(Storage = "_status", Name = "Status", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Status
        {
            get
            {
                return this._status;
            }
            set
            {
                if (((_status == value)
                            == false))
                {
                    this.OnStatusChanging(value);
                    this.SendPropertyChanging();
                    this._status = value;
                    this.SendPropertyChanged("Status");
                    this.OnStatusChanged();
                }
            }
        }

        [Column(Storage = "_targetNation", Name = "TargetNation", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string TargetNation
        {
            get
            {
                return this._targetNation;
            }
            set
            {
                if (((_targetNation == value)
                            == false))
                {
                    this.OnTargetNationChanging(value);
                    this.SendPropertyChanging();
                    this._targetNation = value;
                    this.SendPropertyChanged("TargetNation");
                    this.OnTargetNationChanged();
                }
            }
        }

        [Column(Storage = "_ticker", Name = "Ticker", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Ticker
        {
            get
            {
                return this._ticker;
            }
            set
            {
                if (((_ticker == value)
                            == false))
                {
                    this.OnTickerChanging(value);
                    this.SendPropertyChanging();
                    this._ticker = value;
                    this.SendPropertyChanged("Ticker");
                    this.OnTickerChanged();
                }
            }
        }

        [Column(Storage = "_time", Name = "Time", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Time
        {
            get
            {
                return this._time;
            }
            set
            {
                if (((_time == value)
                            == false))
                {
                    this.OnTimeChanging(value);
                    this.SendPropertyChanging();
                    this._time = value;
                    this.SendPropertyChanged("Time");
                    this.OnTimeChanged();
                }
            }
        }

        [Column(Storage = "_title", Name = "Title", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Title
        {
            get
            {
                return this._title;
            }
            set
            {
                if (((_title == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._title = value;
                    this.SendPropertyChanged("Title");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_sourcingEngine", Name = "SourcingEngine", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourcingEngine
        {
            get
            {
                return this._sourcingEngine;
            }
            set
            {
                if (((_sourcingEngine == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._sourcingEngine = value;
                    this.SendPropertyChanged("SourcingEngine");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_sourceOfArticle", Name = "SourceOfArticle", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceOfArticle
        {
            get
            {
                return this._sourceOfArticle;
            }
            set
            {
                if (((_sourceOfArticle == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._sourceOfArticle = value;
                    this.SendPropertyChanged("SourceOfArticle");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_lastModifiedBy", Name = "LastModifiedBy", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string LastModifiedBy
        {
            get
            {
                return this._lastModifiedBy;
            }
            set
            {
                if (((_lastModifiedBy == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._lastModifiedBy = value;
                    this.SendPropertyChanged("LastModifiedBy");
                    this.OnTitleChanged();
                }
            }
        }

        public string PDF_LocalPath { get; set; }
        public string TXT_LocalPath { get; set; }
        public string MarketCode { get; set; }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        #region ICloneable Members

        public object Clone()
        {
            CHNProcessItem item = new CHNProcessItem();
            item.AssignedTo = this.AssignedTo;
            item.Comments = this.Comments;
            item.Creator = this.Creator;
            item.DealType = this.DealType;
            item.SourceFrom = this.SourceFrom;
            item.EnglishName = this.EnglishName;
            item.ID = this.ID;
            item.IsNew = this.IsNew;
            item.LogNumber = this.LogNumber;
            item.MarketName = this.MarketName;
            item.NumberOfDeals = this.NumberOfDeals;
            item.ScopeType = this.ScopeType;
            item.AnnouceDate = this.AnnouceDate;
            item.Time = this.Time;
            item.SourceLink = this.SourceLink;
            item.Status = this.Status;
            item.TargetNation = this.TargetNation;
            item.Ticker = this.Ticker;
            item.Title = this.Title;
            item.MarketName = this.MarketName;
            item.FilterKeyword = this.FilterKeyword;
            item.ContentKeyword = this.ContentKeyword;
            item.HtmlPath = this.HtmlPath;
            item.CreateDate = this.CreateDate;
            item.LastModify = this.LastModify;

            item.MarketCode = this.MarketCode;
            item.PDF_LocalPath = this.PDF_LocalPath;
            item.TXT_LocalPath = this.TXT_LocalPath;

            item.Content = this.Content;
            item.Url = this.Url;
            item.CaptureDate = this.CaptureDate;
            item.SourcingEngine = this.SourcingEngine;
            item.FinishedDate = this.FinishedDate;

            item.PeComments = this.PeComments;
            item.SourceOfArticle = this.SourceOfArticle;

            return item;
        }

        #endregion
    }

    [Table(Name = "deals.source")]
    public partial class KORProcessItem : ProcessItem, System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged, ICloneable
    {
        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private DateTime _annouceDate;

        private string _assignedTo = string.Empty;

        private string _comments = string.Empty;

        private string _contentKeyword = string.Empty;

        private string _creator = string.Empty;

        private string _dealType = string.Empty;

        private string _englishName = string.Empty;

        private string _filterKeyword = string.Empty;

        private string _htmlPath = string.Empty;

        private int _id;

        private string _isNew = string.Empty;

        private string _logNumber = string.Empty;

        private string _marketName = string.Empty;

        private string _numberOfDeals = string.Empty;

        private string _scopeType = string.Empty;

        private string _sourceFrom = string.Empty;

        private string _sourceLink = string.Empty;

        private string _status = string.Empty;

        private string _targetNation = string.Empty;

        private string _ticker = string.Empty;

        private string _time = string.Empty;

        private string _title = string.Empty;

        private string _sourcingEngine = string.Empty;

        private DateTime _createDate;

        private DateTime _lastModify;

        private DateTime _finishedDate;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnAnnouceDateChanged();

        partial void OnAnnouceDateChanging(DateTime value);

        partial void OnAssignedToChanged();

        partial void OnAssignedToChanging(string value);

        partial void OnCommentsChanged();

        partial void OnCommentsChanging(string value);

        partial void OnContentKeywordChanged();

        partial void OnContentKeywordChanging(string value);

        partial void OnCreatorChanged();

        partial void OnCreatorChanging(string value);

        partial void OnDealTypeChanged();

        partial void OnDealTypeChanging(string value);

        partial void OnEnglishNameChanged();

        partial void OnEnglishNameChanging(string value);

        partial void OnFilterKeywordChanged();

        partial void OnFilterKeywordChanging(string value);

        partial void OnHtmlPathChanged();

        partial void OnHtmlPathChanging(string value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIsNewChanged();

        partial void OnIsNewChanging(string value);

        partial void OnLogNumberChanged();

        partial void OnLogNumberChanging(string value);

        partial void OnMarketNameChanged();

        partial void OnMarketNameChanging(string value);

        partial void OnNumberOfDealsChanged();

        partial void OnNumberOfDealsChanging(string value);

        partial void OnScopeTypeChanged();

        partial void OnScopeTypeChanging(string value);

        partial void OnSourceFromChanged();

        partial void OnSourceFromChanging(string value);

        partial void OnSourceLinkChanged();

        partial void OnSourceLinkChanging(string value);

        partial void OnStatusChanged();

        partial void OnStatusChanging(string value);

        partial void OnTargetNationChanged();

        partial void OnTargetNationChanging(string value);

        partial void OnTickerChanged();

        partial void OnTickerChanging(string value);

        partial void OnTimeChanged();

        partial void OnTimeChanging(string value);

        partial void OnTitleChanged();

        partial void OnTitleChanging(string value);

        partial void OnSourcingEngineChanged();

        partial void OnSourcingEngineChanging(string value);

        partial void OnCreateDateChanged();

        partial void OnCreateDateChanging(DateTime value);

        partial void OnLastModifyChanged();

        partial void OnLastModifyChanging(DateTime value);

        partial void OnFinishedDateChanged();

        partial void OnFinishedDateChanging(DateTime value);
        #endregion


        public KORProcessItem()
        {
            this.OnCreated();
        }

        [Column(Storage = "_createDate", Name = "CreateDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime CreateDate
        {
            get
            {
                return _createDate;
            }
            set
            {
                if (((_createDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._createDate = value;
                    this.SendPropertyChanged("CreateDate");
                    this.OnCreateDateChanged();
                }
            }
        }

        [Column(Storage = "_finishedDate", Name = "FinishedDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime FinishedDate
        {
            get
            {
                return _finishedDate;
            }
            set
            {
                if (((_finishedDate == value)
                            == false))
                {
                    this.OnCreateDateChanging(value);
                    this.SendPropertyChanging();
                    this._finishedDate = value;
                    this.SendPropertyChanged("FinishedDate");
                    this.OnFinishedDateChanged();
                }
            }
        }

        public string CreateDate_
        {
            get
            {
                return CreateDate.ToString();
            }
        }

        public string LastModify_
        {
            get
            {
                return LastModify.ToString();
            }
        }

        public string AnnouceDate_
        {
            get
            {
                return AnnouceDate.ToString();
            }
        }

        [Column(Storage = "_lastModify", Name = "LastModify", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime LastModify
        {
            get
            {
                return _lastModify;
            }
            set
            {
                if (((_lastModify == value)
                            == false))
                {
                    this.OnLastModifyChanging(value);
                    this.SendPropertyChanging();
                    this._lastModify = value;
                    this.SendPropertyChanged("LastModify");
                    this.OnLastModifyChanged();
                }
            }
        }

        [Column(Storage = "_annouceDate", Name = "AnnouceDate", DbType = "datetime", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public DateTime AnnouceDate
        {
            get
            {
                return this._annouceDate;
            }
            set
            {
                if (((_annouceDate == value)
                            == false))
                {
                    this.OnAnnouceDateChanging(value);
                    this.SendPropertyChanging();
                    this._annouceDate = value;
                    this.SendPropertyChanged("AnnouceDate");
                    this.OnAnnouceDateChanged();
                }
            }
        }

        [Column(Storage = "_assignedTo", Name = "AssignedTo", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string AssignedTo
        {
            get
            {
                return this._assignedTo;
            }
            set
            {
                if (((_assignedTo == value)
                            == false))
                {
                    this.OnAssignedToChanging(value);
                    this.SendPropertyChanging();
                    this._assignedTo = value;
                    this.SendPropertyChanged("AssignedTo");
                    this.OnAssignedToChanged();
                }
            }
        }

        [Column(Storage = "_comments", Name = "Comments", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Comments
        {
            get
            {
                return this._comments;
            }
            set
            {
                if (((_comments == value)
                            == false))
                {
                    this.OnCommentsChanging(value);
                    this.SendPropertyChanging();
                    this._comments = value;
                    this.SendPropertyChanged("Comments");
                    this.OnCommentsChanged();
                }
            }
        }

        [Column(Storage = "_contentKeyword", Name = "ContentKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ContentKeyword
        {
            get
            {
                return this._contentKeyword;
            }
            set
            {
                if (((_contentKeyword == value)
                            == false))
                {
                    this.OnContentKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._contentKeyword = value;
                    this.SendPropertyChanged("ContentKeyword");
                    this.OnContentKeywordChanged();
                }
            }
        }

        [Column(Storage = "_creator", Name = "Creator", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Creator
        {
            get
            {
                return this._creator;
            }
            set
            {
                if (((_creator == value)
                            == false))
                {
                    this.OnCreatorChanging(value);
                    this.SendPropertyChanging();
                    this._creator = value;
                    this.SendPropertyChanged("Creator");
                    this.OnCreatorChanged();
                }
            }
        }

        [Column(Storage = "_dealType", Name = "DealType", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string DealType
        {
            get
            {
                return this._dealType;
            }
            set
            {
                if (((_dealType == value)
                            == false))
                {
                    this.OnDealTypeChanging(value);
                    this.SendPropertyChanging();
                    this._dealType = value;
                    this.SendPropertyChanged("DealType");
                    this.OnDealTypeChanged();
                }
            }
        }

        [Column(Storage = "_englishName", Name = "EnglishName", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string EnglishName
        {
            get
            {
                return this._englishName;
            }
            set
            {
                if (((_englishName == value)
                            == false))
                {
                    this.OnEnglishNameChanging(value);
                    this.SendPropertyChanging();
                    this._englishName = value;
                    this.SendPropertyChanged("EnglishName");
                    this.OnEnglishNameChanged();
                }
            }
        }

        [Column(Storage = "_filterKeyword", Name = "FilterKeyword", DbType = "varchar(1000)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string FilterKeyword
        {
            get
            {
                return this._filterKeyword;
            }
            set
            {
                if (((_filterKeyword == value)
                            == false))
                {
                    this.OnFilterKeywordChanging(value);
                    this.SendPropertyChanging();
                    this._filterKeyword = value;
                    this.SendPropertyChanged("FilterKeyword");
                    this.OnFilterKeywordChanged();
                }
            }
        }

        [Column(Storage = "_htmlPath", Name = "HtmlPath", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string HtmlPath
        {
            get
            {
                return this._htmlPath;
            }
            set
            {
                if (((_htmlPath == value)
                            == false))
                {
                    this.OnHtmlPathChanging(value);
                    this.SendPropertyChanging();
                    this._htmlPath = value;
                    this.SendPropertyChanged("HtmlPath");
                    this.OnHtmlPathChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "Id", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_isNew", Name = "IsNew", DbType = "varchar(10)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string IsNew
        {
            get
            {
                return this._isNew;
            }
            set
            {
                if ((_isNew != value))
                {
                    this.OnIsNewChanging(value);
                    this.SendPropertyChanging();
                    this._isNew = value;
                    this.SendPropertyChanged("IsNew");
                    this.OnIsNewChanged();
                }
            }
        }

        [Column(Storage = "_logNumber", Name = "LogNumber", DbType = "varchar(255)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string LogNumber
        {
            get
            {
                return this._logNumber;
            }
            set
            {
                if (((_logNumber == value)
                            == false))
                {
                    this.OnLogNumberChanging(value);
                    this.SendPropertyChanging();
                    this._logNumber = value;
                    this.SendPropertyChanged("LogNumber");
                    this.OnLogNumberChanged();
                }
            }
        }

        [Column(Storage = "_marketName", Name = "MarketName", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string MarketName
        {
            get
            {
                return this._marketName;
            }
            set
            {
                if (((_marketName == value)
                            == false))
                {
                    this.OnMarketNameChanging(value);
                    this.SendPropertyChanging();
                    this._marketName = value;
                    this.SendPropertyChanged("MarketName");
                    this.OnMarketNameChanged();
                }
            }
        }

        [Column(Storage = "_numberOfDeals", Name = "NumberOfDeals", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string NumberOfDeals
        {
            get
            {
                return this._numberOfDeals;
            }
            set
            {
                if ((_numberOfDeals != value))
                {
                    this.OnNumberOfDealsChanging(value);
                    this.SendPropertyChanging();
                    this._numberOfDeals = value;
                    this.SendPropertyChanged("NumberOfDeals");
                    this.OnNumberOfDealsChanged();
                }
            }
        }

        [Column(Storage = "_scopeType", Name = "ScopeType", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string ScopeType
        {
            get
            {
                return this._scopeType;
            }
            set
            {
                if (((_scopeType == value)
                            == false))
                {
                    this.OnScopeTypeChanging(value);
                    this.SendPropertyChanging();
                    this._scopeType = value;
                    this.SendPropertyChanged("ScopeType");
                    this.OnScopeTypeChanged();
                }
            }
        }

        [Column(Storage = "_sourceFrom", Name = "SourceFrom", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceFrom
        {
            get
            {
                return this._sourceFrom;
            }
            set
            {
                if (((_sourceFrom == value)
                            == false))
                {
                    this.OnSourceFromChanging(value);
                    this.SendPropertyChanging();
                    this._sourceFrom = value;
                    this.SendPropertyChanged("SourceFrom");
                    this.OnSourceFromChanged();
                }
            }
        }

        [Column(Storage = "_sourceLink", Name = "SourceLink", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourceLink
        {
            get
            {
                return this._sourceLink;
            }
            set
            {
                if (((_sourceLink == value)
                            == false))
                {
                    this.OnSourceLinkChanging(value);
                    this.SendPropertyChanging();
                    this._sourceLink = value;
                    this.SendPropertyChanged("SourceLink");
                    this.OnSourceLinkChanged();
                }
            }
        }

        [Column(Storage = "_status", Name = "Status", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Status
        {
            get
            {
                return this._status;
            }
            set
            {
                if (((_status == value)
                            == false))
                {
                    this.OnStatusChanging(value);
                    this.SendPropertyChanging();
                    this._status = value;
                    this.SendPropertyChanged("Status");
                    this.OnStatusChanged();
                }
            }
        }

        [Column(Storage = "_targetNation", Name = "TargetNation", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string TargetNation
        {
            get
            {
                return this._targetNation;
            }
            set
            {
                if (((_targetNation == value)
                            == false))
                {
                    this.OnTargetNationChanging(value);
                    this.SendPropertyChanging();
                    this._targetNation = value;
                    this.SendPropertyChanged("TargetNation");
                    this.OnTargetNationChanged();
                }
            }
        }

        [Column(Storage = "_ticker", Name = "Ticker", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Ticker
        {
            get
            {
                return this._ticker;
            }
            set
            {
                if (((_ticker == value)
                            == false))
                {
                    this.OnTickerChanging(value);
                    this.SendPropertyChanging();
                    this._ticker = value;
                    this.SendPropertyChanged("Ticker");
                    this.OnTickerChanged();
                }
            }
        }

        [Column(Storage = "_time", Name = "Time", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Time
        {
            get
            {
                return this._time;
            }
            set
            {
                if (((_time == value)
                            == false))
                {
                    this.OnTimeChanging(value);
                    this.SendPropertyChanging();
                    this._time = value;
                    this.SendPropertyChanged("Time");
                    this.OnTimeChanged();
                }
            }
        }

        [Column(Storage = "_title", Name = "Title", DbType = "varchar(100)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Title
        {
            get
            {
                return this._title;
            }
            set
            {
                if (((_title == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._title = value;
                    this.SendPropertyChanged("Title");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_sourcingEngine", Name = "SourcingEngine", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string SourcingEngine
        {
            get
            {
                return this._sourcingEngine;
            }
            set
            {
                if (((_sourcingEngine == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._sourcingEngine = value;
                    this.SendPropertyChanged("SourcingEngine");
                    this.OnTitleChanged();
                }
            }
        }

        public string PDF_LocalPath { get; set; }
        public string TXT_LocalPath { get; set; }
        public string MarketCode { get; set; }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        #region ICloneable Members

        public object Clone()
        {
            KORProcessItem item = new KORProcessItem();
            item.AssignedTo = this.AssignedTo;
            item.Comments = this.Comments;
            item.Creator = this.Creator;
            item.DealType = this.DealType;
            item.SourceFrom = this.SourceFrom;
            item.EnglishName = this.EnglishName;
            item.ID = this.ID;
            item.IsNew = this.IsNew;
            item.LogNumber = this.LogNumber;
            item.MarketName = this.MarketName;
            item.NumberOfDeals = this.NumberOfDeals;
            item.ScopeType = this.ScopeType;
            item.AnnouceDate = this.AnnouceDate;
            item.Time = this.Time;
            item.SourceLink = this.SourceLink;
            item.Status = this.Status;
            item.TargetNation = this.TargetNation;
            item.Ticker = this.Ticker;
            item.Title = this.Title;
            item.MarketName = this.MarketName;
            item.FilterKeyword = this.FilterKeyword;
            item.ContentKeyword = this.ContentKeyword;
            item.HtmlPath = this.HtmlPath;
            item.CreateDate = this.CreateDate;
            item.LastModify = this.LastModify;

            item.MarketCode = this.MarketCode;
            item.PDF_LocalPath = this.PDF_LocalPath;
            item.TXT_LocalPath = this.TXT_LocalPath;

            item.Content = this.Content;
            item.Url = this.Url;
            item.CaptureDate = this.CaptureDate;
            item.SourcingEngine = this.SourcingEngine;
            item.FinishedDate = this.FinishedDate;
            return item;
        }

        #endregion
    }

    [Table(Name = "deals.statistics")]
    public partial class Statistics : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private string _countryCode;

        private string _date;

        private int _inScope;

        private string _market;

        private int _outOfScope;

        private int _total;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnCountryCodeChanged();

        partial void OnCountryCodeChanging(string value);

        partial void OnDateChanged();

        partial void OnDateChanging(string value);

        partial void OnInScopeChanged();

        partial void OnInScopeChanging(int value);

        partial void OnMarketChanged();

        partial void OnMarketChanging(string value);

        partial void OnOutOfScopeChanged();

        partial void OnOutOfScopeChanging(int value);

        partial void OnTotalChanged();

        partial void OnTotalChanging(int value);
        #endregion


        public Statistics()
        {
            this.OnCreated();
        }

        [Column(Storage = "_countryCode", Name = "CountryCode", IsPrimaryKey = true, DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string CountryCode
        {
            get
            {
                return this._countryCode;
            }
            set
            {
                if (((_countryCode == value)
                            == false))
                {
                    this.OnCountryCodeChanging(value);
                    this.SendPropertyChanging();
                    this._countryCode = value;
                    this.SendPropertyChanged("CountryCode");
                    this.OnCountryCodeChanged();
                }
            }
        }

        [Column(Storage = "_date", Name = "Date", IsPrimaryKey = true, DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Date
        {
            get
            {
                return this._date;
            }
            set
            {
                if (((_date == value)
                            == false))
                {
                    this.OnDateChanging(value);
                    this.SendPropertyChanging();
                    this._date = value;
                    this.SendPropertyChanged("Date");
                    this.OnDateChanged();
                }
            }
        }

        [Column(Storage = "_inScope", Name = "InScope", DbType = "int", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int InScope
        {
            get
            {
                return this._inScope;
            }
            set
            {
                if ((_inScope != value))
                {
                    this.OnInScopeChanging(value);
                    this.SendPropertyChanging();
                    this._inScope = value;
                    this.SendPropertyChanged("InScope");
                    this.OnInScopeChanged();
                }
            }
        }

        [Column(Storage = "_market", Name = "Market", IsPrimaryKey = true, DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Market
        {
            get
            {
                return this._market;
            }
            set
            {
                if (((_market == value)
                            == false))
                {
                    this.OnMarketChanging(value);
                    this.SendPropertyChanging();
                    this._market = value;
                    this.SendPropertyChanged("Market");
                    this.OnMarketChanged();
                }
            }
        }

        [Column(Storage = "_outOfScope", Name = "OutOfScope", DbType = "int", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int OutOfScope
        {
            get
            {
                return this._outOfScope;
            }
            set
            {
                if ((_outOfScope != value))
                {
                    this.OnOutOfScopeChanging(value);
                    this.SendPropertyChanging();
                    this._outOfScope = value;
                    this.SendPropertyChanged("OutOfScope");
                    this.OnOutOfScopeChanged();
                }
            }
        }

        [Column(Storage = "_total", Name = "Total", DbType = "int", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int Total
        {
            get
            {
                return this._total;
            }
            set
            {
                if ((_total != value))
                {
                    this.OnTotalChanging(value);
                    this.SendPropertyChanging();
                    this._total = value;
                    this.SendPropertyChanged("Total");
                    this.OnTotalChanged();
                }
            }
        }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }

    [Table(Name = "asd.configuration")]
    public partial class Configuration : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private string _key;

        private string _value;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnKeyChanged();

        partial void OnKeyChanging(string value);

        partial void OnValueChanged();

        partial void OnValueChanging(string value);
        #endregion


        public Configuration()
        {
            this.OnCreated();
        }

        [Column(Storage = "_key", Name = "Key", DbType = "varchar(50)", IsPrimaryKey = true, AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Key
        {
            get
            {
                return this._key;
            }
            set
            {
                if (((_key == value)
                            == false))
                {
                    this.OnKeyChanging(value);
                    this.SendPropertyChanging();
                    this._key = value;
                    this.SendPropertyChanged("Key");
                    this.OnKeyChanged();
                }
            }
        }

        [Column(Storage = "_value", Name = "Value", DbType = "text", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string Value
        {
            get
            {
                return this._value;
            }
            set
            {
                if (((_value == value)
                            == false))
                {
                    this.OnValueChanging(value);
                    this.SendPropertyChanging();
                    this._value = value;
                    this.SendPropertyChanged("Value");
                    this.OnValueChanged();
                }
            }
        }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }

    [Table(Name = "deals.tickers")]
    public partial class Tickers : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private uint _id;

        private string _tickerA;

        private string _tickerB;

        private string _nameA;

        private string _nameB;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnTickerAChanged();

        partial void OnTickerAChanging(string value);

        partial void OnTickerBChanged();

        partial void OnTickerBChanging(string value);

        partial void OnIDChanged();

        partial void OnIDChanging(uint value);

        partial void OnNameAChanged();

        partial void OnNameAChanging(string value);

        partial void OnNameBChanged();

        partial void OnNameBChanging(string value);
        #endregion


        public Tickers()
        {
            this.OnCreated();
        }

        [Column(Storage = "_id", Name = "Id", DbType = "int unsigned", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public uint ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_tickerA", Name = "TickerA", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string TickerA
        {
            get
            {
                return this._tickerA;
            }
            set
            {
                if (((_tickerA == value)
                            == false))
                {
                    this.OnTickerAChanging(value);
                    this.SendPropertyChanging();
                    this._tickerA = value;
                    this.SendPropertyChanged("TickerA");
                    this.OnTickerAChanged();
                }
            }
        }

        [Column(Storage = "_tickerB", Name = "TickerB", DbType = "varchar(20)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string TickerB
        {
            get
            {
                return this._tickerB;
            }
            set
            {
                if (((_tickerB == value)
                            == false))
                {
                    this.OnTickerBChanging(value);
                    this.SendPropertyChanging();
                    this._tickerB = value;
                    this.SendPropertyChanged("TickerB");
                    this.OnTickerBChanged();
                }
            }
        }

        [Column(Storage = "_nameA", Name = "NameA", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string NameA
        {
            get
            {
                return this._nameA;
            }
            set
            {
                if (((_nameA == value)
                            == false))
                {
                    this.OnNameAChanging(value);
                    this.SendPropertyChanging();
                    this._nameA = value;
                    this.SendPropertyChanged("NameA");
                    this.OnNameAChanged();
                }
            }
        }

        [Column(Storage = "_nameB", Name = "NameB", DbType = "varchar(50)", AutoSync = System.Data.Linq.Mapping.AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public string NameB
        {
            get
            {
                return this._nameB;
            }
            set
            {
                if (((_nameB == value)
                            == false))
                {
                    this.OnNameBChanging(value);
                    this.SendPropertyChanging();
                    this._nameB = value;
                    this.SendPropertyChanged("NameB");
                    this.OnNameBChanged();
                }
            }
        }

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }

    [Table(Name = "deals.ipo_annouce")]
    public partial class IPOAnnOUCe : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private System.Nullable<System.DateTime> _createdDate;

        private int _id;

        private System.Nullable<int> _ipOiD;

        private string _link;

        private System.Nullable<System.DateTime> _revealDate;

        private string _source;

        private string _title;

        private string _type;

        private EntityRef<IPOSource> _iposOurce = new EntityRef<IPOSource>();

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnCreatedDateChanged();

        partial void OnCreatedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIPoiDChanged();

        partial void OnIPoiDChanging(System.Nullable<int> value);

        partial void OnLinkChanged();

        partial void OnLinkChanging(string value);

        partial void OnRevealDateChanged();

        partial void OnRevealDateChanging(System.Nullable<System.DateTime> value);

        partial void OnSourceChanged();

        partial void OnSourceChanging(string value);

        partial void OnTitleChanged();

        partial void OnTitleChanging(string value);

        partial void OnTypeChanged();

        partial void OnTypeChanging(string value);
        #endregion


        public IPOAnnOUCe()
        {
            this.OnCreated();
        }

        [Column(Storage = "_createdDate", Name = "CreatedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> CreatedDate
        {
            get
            {
                return this._createdDate;
            }
            set
            {
                if ((_createdDate != value))
                {
                    this.OnCreatedDateChanging(value);
                    this.SendPropertyChanging();
                    this._createdDate = value;
                    this.SendPropertyChanged("CreatedDate");
                    this.OnCreatedDateChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "ID", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_ipOiD", Name = "IPOID", DbType = "int", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<int> IPoiD
        {
            get
            {
                return this._ipOiD;
            }
            set
            {
                if ((_ipOiD != value))
                {
                    this.OnIPoiDChanging(value);
                    this.SendPropertyChanging();
                    this._ipOiD = value;
                    this.SendPropertyChanged("IPoiD");
                    this.OnIPoiDChanged();
                }
            }
        }

        [Column(Storage = "_link", Name = "Link", DbType = "varchar(1000)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Link
        {
            get
            {
                return this._link;
            }
            set
            {
                if (((_link == value)
                            == false))
                {
                    this.OnLinkChanging(value);
                    this.SendPropertyChanging();
                    this._link = value;
                    this.SendPropertyChanged("Link");
                    this.OnLinkChanged();
                }
            }
        }

        [Column(Storage = "_revealDate", Name = "RevealDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> RevealDate
        {
            get
            {
                return this._revealDate;
            }
            set
            {
                if ((_revealDate != value))
                {
                    this.OnRevealDateChanging(value);
                    this.SendPropertyChanging();
                    this._revealDate = value;
                    this.SendPropertyChanged("RevealDate");
                    this.OnRevealDateChanged();
                }
            }
        }

        [Column(Storage = "_source", Name = "Source", DbType = "varchar(10)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Source
        {
            get
            {
                return this._source;
            }
            set
            {
                if (((_source == value)
                            == false))
                {
                    this.OnSourceChanging(value);
                    this.SendPropertyChanging();
                    this._source = value;
                    this.SendPropertyChanged("Source");
                    this.OnSourceChanged();
                }
            }
        }

        [Column(Storage = "_title", Name = "Title", DbType = "varchar(100)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Title
        {
            get
            {
                return this._title;
            }
            set
            {
                if (((_title == value)
                            == false))
                {
                    this.OnTitleChanging(value);
                    this.SendPropertyChanging();
                    this._title = value;
                    this.SendPropertyChanged("Title");
                    this.OnTitleChanged();
                }
            }
        }

        [Column(Storage = "_type", Name = "Type", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Type
        {
            get
            {
                return this._type;
            }
            set
            {
                if (((_type == value)
                            == false))
                {
                    this.OnTypeChanging(value);
                    this.SendPropertyChanging();
                    this._type = value;
                    this.SendPropertyChanged("Type");
                    this.OnTypeChanged();
                }
            }
        }

        #region Parents
        [Association(Storage = "_iposOurce", OtherKey = "ID", ThisKey = "IPoiD", Name = "FK_IPOID", IsForeignKey = true)]
        [DebuggerNonUserCode()]
        public IPOSource IPOSource
        {
            get
            {
                return this._iposOurce.Entity;
            }
            set
            {
                if (((this._iposOurce.Entity == value)
                            == false))
                {
                    if ((this._iposOurce.Entity != null))
                    {
                        IPOSource previousIPOSource = this._iposOurce.Entity;
                        this._iposOurce.Entity = null;
                        previousIPOSource.IPOAnnOUCe.Remove(this);
                    }
                    this._iposOurce.Entity = value;
                    if ((value != null))
                    {
                        value.IPOAnnOUCe.Add(this);
                        _ipOiD = value.ID;
                    }
                    else
                    {
                        _ipOiD = null;
                    }
                }
            }
        }
        #endregion

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }

    [Table(Name = "deals.ipo_source")]
    public partial class IPOSource : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private string _companyName;

        private System.Nullable<System.DateTime> _createdDate;

        private string _englishName;

        private string _ibstIcker;

        private System.Nullable<System.DateTime> _ibstIckerAddedDate;

        private int _id;

        private string _isIn;

        private System.Nullable<System.DateTime> _isinaDdedDate;

        private System.Nullable<System.DateTime> _listingDate;

        private string _longName;

        private string _market;

        private string _orgID;

        private System.Nullable<System.DateTime> _orgIdaDdedDate;

        private string _permID;

        private System.Nullable<System.DateTime> _permIdaDdedDate;

        private string _piLc;

        private System.Nullable<System.DateTime> _pilcaDdedDate;

        private System.Nullable<int> _processPhase;

        private string _repno;

        private System.Nullable<System.DateTime> _repnoAddedDate;

        private string _riC;

        private System.Nullable<System.DateTime> _ricaDdedDate;

        private string _source;

        private string _ticker;

        private EntitySet<IPOAnnOUCe> _ipoaNnOucE;

        private EntitySet<ReferenceRequest> _referenceRequest;

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnCompanyNameChanged();

        partial void OnCompanyNameChanging(string value);

        partial void OnCreatedDateChanged();

        partial void OnCreatedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnEnglishNameChanged();

        partial void OnEnglishNameChanging(string value);

        partial void OnIbstIckerChanged();

        partial void OnIbstIckerChanging(string value);

        partial void OnIbstIckerAddedDateChanged();

        partial void OnIbstIckerAddedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIsInChanged();

        partial void OnIsInChanging(string value);

        partial void OnIsinaDdedDateChanged();

        partial void OnIsinaDdedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnListingDateChanged();

        partial void OnListingDateChanging(System.Nullable<System.DateTime> value);

        partial void OnLongNameChanged();

        partial void OnLongNameChanging(string value);

        partial void OnMarketChanged();

        partial void OnMarketChanging(string value);

        partial void OnOrgIDChanged();

        partial void OnOrgIDChanging(string value);

        partial void OnOrgIdaDdedDateChanged();

        partial void OnOrgIdaDdedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnPermIDChanged();

        partial void OnPermIDChanging(string value);

        partial void OnPermIdaDdedDateChanged();

        partial void OnPermIdaDdedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnPiLcChanged();

        partial void OnPiLcChanging(string value);

        partial void OnPilcaDdedDateChanged();

        partial void OnPilcaDdedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnProcessPhaseChanged();

        partial void OnProcessPhaseChanging(System.Nullable<int> value);

        partial void OnRepnoChanged();

        partial void OnRepnoChanging(string value);

        partial void OnRepnoAddedDateChanged();

        partial void OnRepnoAddedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnRiCChanged();

        partial void OnRiCChanging(string value);

        partial void OnRicaDdedDateChanged();

        partial void OnRicaDdedDateChanging(System.Nullable<System.DateTime> value);

        partial void OnSourceChanged();

        partial void OnSourceChanging(string value);

        partial void OnTickerChanged();

        partial void OnTickerChanging(string value);
        #endregion


        public IPOSource()
        {
            _ipoaNnOucE = new EntitySet<IPOAnnOUCe>(new Action<IPOAnnOUCe>(this.IPOAnnOUCe_Attach), new Action<IPOAnnOUCe>(this.IPOAnnOUCe_Detach));
            _referenceRequest = new EntitySet<ReferenceRequest>(new Action<ReferenceRequest>(this.ReferenceRequest_Attach), new Action<ReferenceRequest>(this.ReferenceRequest_Detach));
            this.OnCreated();
        }

        [Column(Storage = "_companyName", Name = "CompanyName", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string CompanyName
        {
            get
            {
                return this._companyName;
            }
            set
            {
                if (((_companyName == value)
                            == false))
                {
                    this.OnCompanyNameChanging(value);
                    this.SendPropertyChanging();
                    this._companyName = value;
                    this.SendPropertyChanged("CompanyName");
                    this.OnCompanyNameChanged();
                }
            }
        }

        [Column(Storage = "_createdDate", Name = "CreatedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> CreatedDate
        {
            get
            {
                return this._createdDate;
            }
            set
            {
                if ((_createdDate != value))
                {
                    this.OnCreatedDateChanging(value);
                    this.SendPropertyChanging();
                    this._createdDate = value;
                    this.SendPropertyChanged("CreatedDate");
                    this.OnCreatedDateChanged();
                }
            }
        }

        [Column(Storage = "_englishName", Name = "EnglishName", DbType = "varchar(500)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string EnglishName
        {
            get
            {
                return this._englishName;
            }
            set
            {
                if (((_englishName == value)
                            == false))
                {
                    this.OnEnglishNameChanging(value);
                    this.SendPropertyChanging();
                    this._englishName = value;
                    this.SendPropertyChanged("EnglishName");
                    this.OnEnglishNameChanged();
                }
            }
        }

        [Column(Storage = "_ibstIcker", Name = "IBSTicker", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string IbstIcker
        {
            get
            {
                return this._ibstIcker;
            }
            set
            {
                if (((_ibstIcker == value)
                            == false))
                {
                    this.OnIbstIckerChanging(value);
                    this.SendPropertyChanging();
                    this._ibstIcker = value;
                    this.SendPropertyChanged("IbstIcker");
                    this.OnIbstIckerChanged();
                }
            }
        }

        [Column(Storage = "_ibstIckerAddedDate", Name = "IBSTickerAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> IbstIckerAddedDate
        {
            get
            {
                return this._ibstIckerAddedDate;
            }
            set
            {
                if ((_ibstIckerAddedDate != value))
                {
                    this.OnIbstIckerAddedDateChanging(value);
                    this.SendPropertyChanging();
                    this._ibstIckerAddedDate = value;
                    this.SendPropertyChanged("IbstIckerAddedDate");
                    this.OnIbstIckerAddedDateChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "ID", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_isIn", Name = "ISIN", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string IsIn
        {
            get
            {
                return this._isIn;
            }
            set
            {
                if (((_isIn == value)
                            == false))
                {
                    this.OnIsInChanging(value);
                    this.SendPropertyChanging();
                    this._isIn = value;
                    this.SendPropertyChanged("IsIn");
                    this.OnIsInChanged();
                }
            }
        }

        [Column(Storage = "_isinaDdedDate", Name = "ISINAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> IsinaDdedDate
        {
            get
            {
                return this._isinaDdedDate;
            }
            set
            {
                if ((_isinaDdedDate != value))
                {
                    this.OnIsinaDdedDateChanging(value);
                    this.SendPropertyChanging();
                    this._isinaDdedDate = value;
                    this.SendPropertyChanged("IsinaDdedDate");
                    this.OnIsinaDdedDateChanged();
                }
            }
        }

        [Column(Storage = "_listingDate", Name = "ListingDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> ListingDate
        {
            get
            {
                return this._listingDate;
            }
            set
            {
                if ((_listingDate != value))
                {
                    this.OnListingDateChanging(value);
                    this.SendPropertyChanging();
                    this._listingDate = value;
                    this.SendPropertyChanged("ListingDate");
                    this.OnListingDateChanged();
                }
            }
        }

        [Column(Storage = "_longName", Name = "LongName", DbType = "varchar(50)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string LongName
        {
            get
            {
                return this._longName;
            }
            set
            {
                if (((_longName == value)
                            == false))
                {
                    this.OnLongNameChanging(value);
                    this.SendPropertyChanging();
                    this._longName = value;
                    this.SendPropertyChanged("LongName");
                    this.OnLongNameChanged();
                }
            }
        }

        [Column(Storage = "_market", Name = "Market", DbType = "varchar(10)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Market
        {
            get
            {
                return this._market;
            }
            set
            {
                if (((_market == value)
                            == false))
                {
                    this.OnMarketChanging(value);
                    this.SendPropertyChanging();
                    this._market = value;
                    this.SendPropertyChanged("Market");
                    this.OnMarketChanged();
                }
            }
        }

        [Column(Storage = "_orgID", Name = "OrgID", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string OrgID
        {
            get
            {
                return this._orgID;
            }
            set
            {
                if (((_orgID == value)
                            == false))
                {
                    this.OnOrgIDChanging(value);
                    this.SendPropertyChanging();
                    this._orgID = value;
                    this.SendPropertyChanged("OrgID");
                    this.OnOrgIDChanged();
                }
            }
        }

        [Column(Storage = "_orgIdaDdedDate", Name = "OrgIDAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> OrgIdaDdedDate
        {
            get
            {
                return this._orgIdaDdedDate;
            }
            set
            {
                if ((_orgIdaDdedDate != value))
                {
                    this.OnOrgIdaDdedDateChanging(value);
                    this.SendPropertyChanging();
                    this._orgIdaDdedDate = value;
                    this.SendPropertyChanged("OrgIdaDdedDate");
                    this.OnOrgIdaDdedDateChanged();
                }
            }
        }

        [Column(Storage = "_permID", Name = "PermID", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string PermID
        {
            get
            {
                return this._permID;
            }
            set
            {
                if (((_permID == value)
                            == false))
                {
                    this.OnPermIDChanging(value);
                    this.SendPropertyChanging();
                    this._permID = value;
                    this.SendPropertyChanged("PermID");
                    this.OnPermIDChanged();
                }
            }
        }

        [Column(Storage = "_permIdaDdedDate", Name = "PermIDAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> PermIdaDdedDate
        {
            get
            {
                return this._permIdaDdedDate;
            }
            set
            {
                if ((_permIdaDdedDate != value))
                {
                    this.OnPermIdaDdedDateChanging(value);
                    this.SendPropertyChanging();
                    this._permIdaDdedDate = value;
                    this.SendPropertyChanged("PermIdaDdedDate");
                    this.OnPermIdaDdedDateChanged();
                }
            }
        }

        [Column(Storage = "_piLc", Name = "PILC", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string PiLc
        {
            get
            {
                return this._piLc;
            }
            set
            {
                if (((_piLc == value)
                            == false))
                {
                    this.OnPiLcChanging(value);
                    this.SendPropertyChanging();
                    this._piLc = value;
                    this.SendPropertyChanged("PiLc");
                    this.OnPiLcChanged();
                }
            }
        }

        [Column(Storage = "_pilcaDdedDate", Name = "PILCAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> PilcaDdedDate
        {
            get
            {
                return this._pilcaDdedDate;
            }
            set
            {
                if ((_pilcaDdedDate != value))
                {
                    this.OnPilcaDdedDateChanging(value);
                    this.SendPropertyChanging();
                    this._pilcaDdedDate = value;
                    this.SendPropertyChanged("PilcaDdedDate");
                    this.OnPilcaDdedDateChanged();
                }
            }
        }

        [Column(Storage = "_processPhase", Name = "ProcessPhase", DbType = "int(2)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<int> ProcessPhase
        {
            get
            {
                return this._processPhase;
            }
            set
            {
                if ((_processPhase != value))
                {
                    this.OnProcessPhaseChanging(value);
                    this.SendPropertyChanging();
                    this._processPhase = value;
                    this.SendPropertyChanged("ProcessPhase");
                    this.OnProcessPhaseChanged();
                }
            }
        }

        [Column(Storage = "_repno", Name = "Repno", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Repno
        {
            get
            {
                return this._repno;
            }
            set
            {
                if (((_repno == value)
                            == false))
                {
                    this.OnRepnoChanging(value);
                    this.SendPropertyChanging();
                    this._repno = value;
                    this.SendPropertyChanged("Repno");
                    this.OnRepnoChanged();
                }
            }
        }

        [Column(Storage = "_repnoAddedDate", Name = "RepnoAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> RepnoAddedDate
        {
            get
            {
                return this._repnoAddedDate;
            }
            set
            {
                if ((_repnoAddedDate != value))
                {
                    this.OnRepnoAddedDateChanging(value);
                    this.SendPropertyChanging();
                    this._repnoAddedDate = value;
                    this.SendPropertyChanged("RepnoAddedDate");
                    this.OnRepnoAddedDateChanged();
                }
            }
        }

        [Column(Storage = "_riC", Name = "RIC", DbType = "varchar(10)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string RiC
        {
            get
            {
                return this._riC;
            }
            set
            {
                if (((_riC == value)
                            == false))
                {
                    this.OnRiCChanging(value);
                    this.SendPropertyChanging();
                    this._riC = value;
                    this.SendPropertyChanged("RiC");
                    this.OnRiCChanged();
                }
            }
        }

        [Column(Storage = "_ricaDdedDate", Name = "RICAddedDate", DbType = "datetime", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<System.DateTime> RicaDdedDate
        {
            get
            {
                return this._ricaDdedDate;
            }
            set
            {
                if ((_ricaDdedDate != value))
                {
                    this.OnRicaDdedDateChanging(value);
                    this.SendPropertyChanging();
                    this._ricaDdedDate = value;
                    this.SendPropertyChanged("RicaDdedDate");
                    this.OnRicaDdedDateChanged();
                }
            }
        }

        [Column(Storage = "_source", Name = "Source", DbType = "varchar(10)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Source
        {
            get
            {
                return this._source;
            }
            set
            {
                if (((_source == value)
                            == false))
                {
                    this.OnSourceChanging(value);
                    this.SendPropertyChanging();
                    this._source = value;
                    this.SendPropertyChanged("Source");
                    this.OnSourceChanged();
                }
            }
        }

        [Column(Storage = "_ticker", Name = "Ticker", DbType = "varchar(10)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Ticker
        {
            get
            {
                return this._ticker;
            }
            set
            {
                if (((_ticker == value)
                            == false))
                {
                    this.OnTickerChanging(value);
                    this.SendPropertyChanging();
                    this._ticker = value;
                    this.SendPropertyChanged("Ticker");
                    this.OnTickerChanged();
                }
            }
        }

        #region Children
        [Association(Storage = "_ipoaNnOucE", OtherKey = "IPoiD", ThisKey = "ID", Name = "FK_IPOID")]
        [DebuggerNonUserCode()]
        public EntitySet<IPOAnnOUCe> IPOAnnOUCe
        {
            get
            {
                return this._ipoaNnOucE;
            }
            set
            {
                this._ipoaNnOucE = value;
            }
        }

        [Association(Storage = "_referenceRequest", OtherKey = "IPoiD", ThisKey = "ID", Name = "FK_IPO_ID")]
        [DebuggerNonUserCode()]
        public EntitySet<ReferenceRequest> ReferenceRequest
        {
            get
            {
                return this._referenceRequest;
            }
            set
            {
                this._referenceRequest = value;
            }
        }
        #endregion

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        #region Attachment handlers
        private void IPOAnnOUCe_Attach(IPOAnnOUCe entity)
        {
            this.SendPropertyChanging();
            entity.IPOSource = this;
        }

        private void IPOAnnOUCe_Detach(IPOAnnOUCe entity)
        {
            this.SendPropertyChanging();
            entity.IPOSource = null;
        }

        private void ReferenceRequest_Attach(ReferenceRequest entity)
        {
            this.SendPropertyChanging();
            entity.IPOSource = this;
        }

        private void ReferenceRequest_Detach(ReferenceRequest entity)
        {
            this.SendPropertyChanging();
            entity.IPOSource = null;
        }
        #endregion
    }

    [Table(Name = "deals.reference_request")]
    public partial class ReferenceRequest : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private string _contentSet;

        private int _id;

        private System.Nullable<int> _ipOiD;

        private string _requestType;

        private string _status;

        private EntitySet<ReferenceRequestDetail> _referenceRequestDetail;

        private EntityRef<IPOSource> _iposOurce = new EntityRef<IPOSource>();

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnContentSetChanged();

        partial void OnContentSetChanging(string value);

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnIPoiDChanged();

        partial void OnIPoiDChanging(System.Nullable<int> value);

        partial void OnRequestTypeChanged();

        partial void OnRequestTypeChanging(string value);

        partial void OnStatusChanged();

        partial void OnStatusChanging(string value);
        #endregion


        public ReferenceRequest()
        {
            _referenceRequestDetail = new EntitySet<ReferenceRequestDetail>(new Action<ReferenceRequestDetail>(this.ReferenceRequestDetail_Attach), new Action<ReferenceRequestDetail>(this.ReferenceRequestDetail_Detach));
            this.OnCreated();
        }

        [Column(Storage = "_contentSet", Name = "ContentSet", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string ContentSet
        {
            get
            {
                return this._contentSet;
            }
            set
            {
                if (((_contentSet == value)
                            == false))
                {
                    this.OnContentSetChanging(value);
                    this.SendPropertyChanging();
                    this._contentSet = value;
                    this.SendPropertyChanged("ContentSet");
                    this.OnContentSetChanged();
                }
            }
        }

        [Column(Storage = "_id", Name = "ID", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_ipOiD", Name = "IPOID", DbType = "int", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<int> IPoiD
        {
            get
            {
                return this._ipOiD;
            }
            set
            {
                if ((_ipOiD != value))
                {
                    this.OnIPoiDChanging(value);
                    this.SendPropertyChanging();
                    this._ipOiD = value;
                    this.SendPropertyChanged("IPoiD");
                    this.OnIPoiDChanged();
                }
            }
        }

        [Column(Storage = "_requestType", Name = "RequestType", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string RequestType
        {
            get
            {
                return this._requestType;
            }
            set
            {
                if (((_requestType == value)
                            == false))
                {
                    this.OnRequestTypeChanging(value);
                    this.SendPropertyChanging();
                    this._requestType = value;
                    this.SendPropertyChanged("RequestType");
                    this.OnRequestTypeChanged();
                }
            }
        }

        [Column(Storage = "_status", Name = "Status", DbType = "varchar(20)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Status
        {
            get
            {
                return this._status;
            }
            set
            {
                if (((_status == value)
                            == false))
                {
                    this.OnStatusChanging(value);
                    this.SendPropertyChanging();
                    this._status = value;
                    this.SendPropertyChanged("Status");
                    this.OnStatusChanged();
                }
            }
        }

        #region Children
        [Association(Storage = "_referenceRequestDetail", OtherKey = "RequestID", ThisKey = "ID", Name = "FK_RequestID")]
        [DebuggerNonUserCode()]
        public EntitySet<ReferenceRequestDetail> ReferenceRequestDetail
        {
            get
            {
                return this._referenceRequestDetail;
            }
            set
            {
                this._referenceRequestDetail = value;
            }
        }
        #endregion

        #region Parents
        [Association(Storage = "_iposOurce", OtherKey = "ID", ThisKey = "IPoiD", Name = "FK_IPO_ID", IsForeignKey = true)]
        [DebuggerNonUserCode()]
        public IPOSource IPOSource
        {
            get
            {
                return this._iposOurce.Entity;
            }
            set
            {
                if (((this._iposOurce.Entity == value)
                            == false))
                {
                    if ((this._iposOurce.Entity != null))
                    {
                        IPOSource previousIPOSource = this._iposOurce.Entity;
                        this._iposOurce.Entity = null;
                        previousIPOSource.ReferenceRequest.Remove(this);
                    }
                    this._iposOurce.Entity = value;
                    if ((value != null))
                    {
                        value.ReferenceRequest.Add(this);
                        _ipOiD = value.ID;
                    }
                    else
                    {
                        _ipOiD = null;
                    }
                }
            }
        }
        #endregion

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }

        #region Attachment handlers
        private void ReferenceRequestDetail_Attach(ReferenceRequestDetail entity)
        {
            this.SendPropertyChanging();
            entity.ReferenceRequest = this;
        }

        private void ReferenceRequestDetail_Detach(ReferenceRequestDetail entity)
        {
            this.SendPropertyChanging();
            entity.ReferenceRequest = null;
        }
        #endregion
    }

    [Table(Name = "deals.reference_request_detail")]
    public partial class ReferenceRequestDetail : System.ComponentModel.INotifyPropertyChanging, System.ComponentModel.INotifyPropertyChanged
    {

        private static System.ComponentModel.PropertyChangingEventArgs emptyChangingEventArgs = new System.ComponentModel.PropertyChangingEventArgs("");

        private int _id;

        private string _key;

        private System.Nullable<int> _requestID;

        private string _type;

        private string _value;

        private EntityRef<ReferenceRequest> _referenceRequest = new EntityRef<ReferenceRequest>();

        #region Extensibility Method Declarations
        partial void OnCreated();

        partial void OnIDChanged();

        partial void OnIDChanging(int value);

        partial void OnKeyChanged();

        partial void OnKeyChanging(string value);

        partial void OnRequestIDChanged();

        partial void OnRequestIDChanging(System.Nullable<int> value);

        partial void OnTypeChanged();

        partial void OnTypeChanging(string value);

        partial void OnValueChanged();

        partial void OnValueChanging(string value);
        #endregion


        public ReferenceRequestDetail()
        {
            this.OnCreated();
        }

        [Column(Storage = "_id", Name = "ID", DbType = "int", IsPrimaryKey = true, IsDbGenerated = true, AutoSync = AutoSync.Never, CanBeNull = false)]
        [DebuggerNonUserCode()]
        public int ID
        {
            get
            {
                return this._id;
            }
            set
            {
                if ((_id != value))
                {
                    this.OnIDChanging(value);
                    this.SendPropertyChanging();
                    this._id = value;
                    this.SendPropertyChanged("ID");
                    this.OnIDChanged();
                }
            }
        }

        [Column(Storage = "_key", Name = "Key", DbType = "varchar(50)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Key
        {
            get
            {
                return this._key;
            }
            set
            {
                if (((_key == value)
                            == false))
                {
                    this.OnKeyChanging(value);
                    this.SendPropertyChanging();
                    this._key = value;
                    this.SendPropertyChanged("Key");
                    this.OnKeyChanged();
                }
            }
        }

        [Column(Storage = "_requestID", Name = "RequestID", DbType = "int", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public System.Nullable<int> RequestID
        {
            get
            {
                return this._requestID;
            }
            set
            {
                if ((_requestID != value))
                {
                    if (_referenceRequest.HasLoadedOrAssignedValue)
                    {
                        throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
                    }
                    this.OnRequestIDChanging(value);
                    this.SendPropertyChanging();
                    this._requestID = value;
                    this.SendPropertyChanged("RequestID");
                    this.OnRequestIDChanged();
                }
            }
        }

        [Column(Storage = "_type", Name = "Type", DbType = "varchar(50)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Type
        {
            get
            {
                return this._type;
            }
            set
            {
                if (((_type == value)
                            == false))
                {
                    this.OnTypeChanging(value);
                    this.SendPropertyChanging();
                    this._type = value;
                    this.SendPropertyChanged("Type");
                    this.OnTypeChanged();
                }
            }
        }

        [Column(Storage = "_value", Name = "Value", DbType = "varchar(1000)", AutoSync = AutoSync.Never)]
        [DebuggerNonUserCode()]
        public string Value
        {
            get
            {
                return this._value;
            }
            set
            {
                if (((_value == value)
                            == false))
                {
                    this.OnValueChanging(value);
                    this.SendPropertyChanging();
                    this._value = value;
                    this.SendPropertyChanged("Value");
                    this.OnValueChanged();
                }
            }
        }

        #region Parents
        [Association(Storage = "_referenceRequest", OtherKey = "ID", ThisKey = "RequestID", Name = "FK_RequestID", IsForeignKey = true)]
        [DebuggerNonUserCode()]
        public ReferenceRequest ReferenceRequest
        {
            get
            {
                return this._referenceRequest.Entity;
            }
            set
            {
                if (((this._referenceRequest.Entity == value)
                            == false))
                {
                    if ((this._referenceRequest.Entity != null))
                    {
                        ReferenceRequest previousReferenceRequest = this._referenceRequest.Entity;
                        this._referenceRequest.Entity = null;
                        previousReferenceRequest.ReferenceRequestDetail.Remove(this);
                    }
                    this._referenceRequest.Entity = value;
                    if ((value != null))
                    {
                        value.ReferenceRequestDetail.Add(this);
                        _requestID = value.ID;
                    }
                    else
                    {
                        _requestID = null;
                    }
                }
            }
        }
        #endregion

        public event System.ComponentModel.PropertyChangingEventHandler PropertyChanging;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        protected virtual void SendPropertyChanging()
        {
            System.ComponentModel.PropertyChangingEventHandler h = this.PropertyChanging;
            if ((h != null))
            {
                h(this, emptyChangingEventArgs);
            }
        }

        protected virtual void SendPropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler h = this.PropertyChanged;
            if ((h != null))
            {
                h(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
