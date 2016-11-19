﻿#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace KumIn_WPF
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="Kumin")]
	public partial class DataClasses1DataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    #endregion
		
		public DataClasses1DataContext() : 
				base(global::KumIn_WPF.Properties.Settings.Default.KuminConnectionString, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public DataClasses1DataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<FStudentTable> FStudentTables
		{
			get
			{
				return this.GetTable<FStudentTable>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.FStudentTable")]
	public partial class FStudentTable
	{
		
		private string _LastName;
		
		private string _FirstName;
		
		private string _Barcode;
		
		private string _RealEmail;
		
		private string _Phone1;
		
		private string _Carrier1;
		
		private string _Phone2;
		
		private string _Carrier2;
		
		public FStudentTable()
		{
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_LastName", DbType="NVarChar(30) NOT NULL", CanBeNull=false)]
		public string LastName
		{
			get
			{
				return this._LastName;
			}
			set
			{
				if ((this._LastName != value))
				{
					this._LastName = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_FirstName", DbType="NVarChar(30) NOT NULL", CanBeNull=false)]
		public string FirstName
		{
			get
			{
				return this._FirstName;
			}
			set
			{
				if ((this._FirstName != value))
				{
					this._FirstName = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Barcode", DbType="VarChar(200)")]
		public string Barcode
		{
			get
			{
				return this._Barcode;
			}
			set
			{
				if ((this._Barcode != value))
				{
					this._Barcode = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_RealEmail", DbType="VarChar(100)")]
		public string RealEmail
		{
			get
			{
				return this._RealEmail;
			}
			set
			{
				if ((this._RealEmail != value))
				{
					this._RealEmail = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Phone1", DbType="VarChar(100)")]
		public string Phone1
		{
			get
			{
				return this._Phone1;
			}
			set
			{
				if ((this._Phone1 != value))
				{
					this._Phone1 = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Carrier1", DbType="VarChar(100)")]
		public string Carrier1
		{
			get
			{
				return this._Carrier1;
			}
			set
			{
				if ((this._Carrier1 != value))
				{
					this._Carrier1 = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Phone2", DbType="VarChar(100)")]
		public string Phone2
		{
			get
			{
				return this._Phone2;
			}
			set
			{
				if ((this._Phone2 != value))
				{
					this._Phone2 = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Carrier2", DbType="VarChar(100)")]
		public string Carrier2
		{
			get
			{
				return this._Carrier2;
			}
			set
			{
				if ((this._Carrier2 != value))
				{
					this._Carrier2 = value;
				}
			}
		}
	}
}
#pragma warning restore 1591
