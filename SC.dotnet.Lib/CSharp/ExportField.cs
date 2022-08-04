using System;

namespace SC.dotnet.Lib.CSharp
{
    public class ExportField
    {
        #region Members
        /// <summary>
        /// 
        /// </summary>
        private Type _type;
        public Type Type
        {
            get { return _type; }
            set
            {
                _type = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private string _propertyName;
        public string PropertyName
        {
            get { return _propertyName; }
            set
            {
                _propertyName = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private string _displayName;
        public string DisplayName
        {
            get { return _displayName; }
            set
            {
                _displayName = value;
            }
        }

        private string _dateTimeFormat;
        public string DateTimeFormat
        {
            get { return _dateTimeFormat; }
            set
            {
                _dateTimeFormat = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private bool _isPercent;
        public bool IsPercent
        {
            get { return _isPercent; }
            set
            {
                _isPercent = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private bool _isMoney;
        public bool IsMoney
        {
            get { return _isMoney; }
            set
            {
                _isMoney = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Length { get; set; }


        /// <summary>
        /// 
        /// </summary>
        private bool _isRate;
        public bool IsRate
        {
            get { return _isRate; }
            set
            {
                _isRate = value;
            }
        }
        #endregion        
    }
}
