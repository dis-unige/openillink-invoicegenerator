using CsvHelper;
using CsvHelper.Configuration;
using System;

namespace Openillink.InvoiceGenerator.Models
{
    public class OrderMap : ClassMap<Order>
    {
        public OrderMap()
        {
            Map(o => o.IlLinkId).Name("illinkid");
            Map(o => o.Stade).Name("stade");
            Map(o => o.Localisation).Name("localisation");
            Map(o => o.LieuRetrait).Name("lieuretrait");
            Map(o => o.Sid).Name("sid");
            Map(o => o.Pid).Name("pid");
            Map(o => o.OrderDate).Name("date");
            Map(o => o.SendDate).Name("envoye");
            Map(o => o.BillingDate).Name("facture");
            Map(o => o.RenewDate).Name("renouveler");
            Map(o => o.Price).Name("prix").ConvertUsing(ConvertPrice);
            Map(o => o.Prepaye).Name("prepaye");
            Map(o => o.Ref).Name("ref");
            Map(o => o.Arrivee).Name("arrivee");
            Map(o => o.Name).Name("nom");
            Map(o => o.FirstName).Name("prenom");
            Map(o => o.Service).Name("service");
            Map(o => o.CgrA).Name("cgra");
            Map(o => o.CgrB).Name("cgrb");
            Map(o => o.InvoiceGroup).Name("invoicegroup").Validate(NotEmpty);
            Map(o => o.InvoiceAccount).Name("invoiceaccount").Validate(NotEmpty);
            Map(o => o.InvoiceOrigin).Name("invoiceorigin");
            Map(o => o.EMail).Name("mail");
            Map(o => o.Telephone).Name("tel");
            Map(o => o.Adresse).Name("adresse");
            Map(o => o.CodePostale).Name("code_postal");
            Map(o => o.Localite).Name("localite");
            Map(o => o.DocumentType).Name("type_doc");
            Map(o => o.Urgent).Name("urgent");
            Map(o => o.Envoi_Par).Name("envoi_par");
            Map(o => o.Title).Name("titre_periodique");
            Map(o => o.Year).Name("annee");
            Map(o => o.Volume).Name("volume");
            Map(o => o.Number).Name("numero");
            Map(o => o.Supplement).Name("supplement");
            Map(o => o.Pages).Name("pages");
            Map(o => o.ArticleTitle).Name("titre_article");
            Map(o => o.Authors).Name("auteurs");
            Map(o => o.Edition).Name("edition");
            Map(o => o.Isbn).Name("isbn");
            Map(o => o.Issn).Name("issn");
            Map(o => o.Eissn).Name("eissn");
            Map(o => o.Doi).Name("doi");
            Map(o => o.Uid).Name("uid");
            Map(o => o.Remarks).Name("remarques");
            Map(o => o.RemarksPublic).Name("remarquespub");
            Map(o => o.RemarksUser).Name("remarquesuser");
            Map(o => o.History).Name("historique");
            Map(o => o.InputBy).Name("saisie_par");
            Map(o => o.Library).Name("bibliotheque");
            Map(o => o.RefInterBib).Name("refinterbib");
            Map(o => o.PmId).Name("PMID");
            Map(o => o.IpAddress).Name("ip");
            Map(o => o.Referer).Name("referer");
        }

        private bool NotEmpty(string arg)
        {
            return !string.IsNullOrEmpty(arg?.Trim());
        }

        private int? ConvertPrice(IReaderRow row)
        {
            var raw = row.GetField<string>("prix");
            return string.IsNullOrEmpty(raw) 
                        ? (int?)null
                        : int.Parse(raw.Split(new[] { ".", "," }, StringSplitOptions.RemoveEmptyEntries)[0]);
        }
    }
}
