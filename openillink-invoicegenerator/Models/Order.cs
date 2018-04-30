using System;

namespace Openillink.InvoiceGenerator.Models
{
    public class Order
    {
        public long IlLinkId { get; set; }

        public string Stade { get; set; }

        public string Localisation { get; set; }

        public string LieuRetrait { get; set; }

        public string Sid { get; set; }

        public string Pid { get; set; }

        public DateTime OrderDate { get; set; }

        public DateTime? SendDate { get; set; }

        public DateTime? BillingDate { get; set; }

        public DateTime? RenewDate { get; set; }

        public int? Price { get; set; }

        public string Prepaye { get; set; }

        public string Ref { get; set; }

        public string Arrivee { get; set; }

        public string Name { get; set; }

        public string FirstName { get; set; }

        public string Service { get; set; }

        public string CgrA { get; set; }

        public string CgrB { get; set; }

        public string InvoiceGroup { get; set; }

        public string InvoiceAccount { get; set; }

        public string InvoiceOrigin { get; set; }

        public string EMail { get; set; }

        public string Telephone { get; set; }

        public string Adresse { get; set; }

        public string CodePostale { get; set; }

        public string Localite { get; set; }

        public string DocumentType { get; set; }

        public int? Urgent { get; set; }

        public string Envoi_Par { get; set; }

        public string Title { get; set; }

        public string Year { get; set; }

        public string Volume { get; set; }

        public string Number { get; set; }

        public string Supplement { get; set; }

        public string Pages { get; set; }

        public string ArticleTitle { get; set; }

        public string Authors { get; set; }

        public string Edition { get; set; }

        public string Isbn { get; set; }

        public string Issn { get; set; }

        public string Eissn { get; set; }

        public string Doi { get; set; }

        public string Uid { get; set; }

        public string Remarks { get; set; }

        public string RemarksPublic { get; set; }

        public string RemarksUser { get; set; }

        public string History { get; set; }

        public string InputBy { get; set; }

        public string Library { get; set; }

        public string RefInterBib { get; set; }

        public string PmId { get; set; }

        public string IpAddress { get; set; }

        public string Referer { get; set; }
    }
}
