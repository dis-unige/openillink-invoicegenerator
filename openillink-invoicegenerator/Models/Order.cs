using FileHelpers;
using System;
using Openillink.InvoiceGenerator.Helpers;

namespace Openillink.InvoiceGenerator.Models
{
    [DelimitedRecord(";")]
    [IgnoreFirst(3)]
    [IgnoreEmptyLines]
    public class Order
    {
        [FieldQuoted]
        public long IlLinkId;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Stade;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Localisation;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string LieuRetrait;

        [FieldValueDiscarded]
        public string Sid;

        [FieldValueDiscarded]
        public string Pid;

        [FieldQuoted]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime OrderDate;

        [FieldQuoted]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime? SendDate;

        [FieldQuoted]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime? BillingDate;

        [FieldQuoted]
        [FieldConverter(ConverterKind.Date, "yyyy-MM-dd")]
        public DateTime? RenewDate;

        [FieldQuoted]
        public int? Price;

        [FieldValueDiscarded]
        public string Prepaye;

        [FieldValueDiscarded]
        public string Ref;

        [FieldValueDiscarded]
        public string Arrivee;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Name;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string FirstName;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Service;

        [FieldValueDiscarded]
        public string CgrA;

        [FieldValueDiscarded]
        public string CgrB;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        [FieldNotEmpty]
        public string InvoiceGroup;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        [FieldNotEmpty]
        public string InvoiceAccount;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string InvoiceOrigin;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string EMail;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Telephone;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Adresse;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string CodePostale;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Localite;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string DocumentType;

        [FieldQuoted]
        public int? Urgent;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Envoi_Par;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Title;

        [FieldQuoted]
        public string Year;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Volume;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Number;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Supplement;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Pages;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string ArticleTitle;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Authors;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Edition;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Isbn;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Issn;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Eissn;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Doi;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Uid;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Remarks;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string RemarksPublic;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string RemarksUser;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string History;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string InputBy;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string Library;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string RefInterBib;

        [FieldQuoted]
        [FieldConverter(typeof(StringConverter))]
        public string PmId;

        [FieldValueDiscarded]
        public string IpAddress;

        [FieldValueDiscarded]
        public string Referer;
    }
}
