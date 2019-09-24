using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SAPbouiCOM.Framework;
using SAPbobsCOM;

namespace CostCenterOutgoing
{
    [FormAttribute("141", "APInvoice.b1f")]
    class APInvoice : SystemFormBase
    {
        public APInvoice()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }



        private void OnCustomInitialize()
        {

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            if (pVal.ActionSuccess)
            {
                int docEntry = 0;
                string xmlObjectKey = pVal.ObjectKey;
                XElement xmlnew = XElement.Parse(xmlObjectKey);
                XElement xElement = xmlnew.Element("DocEntry");
                if (xElement != null)
                {
                    docEntry = int.Parse(xElement.Value);
                }

                Documents apInvoice = (Documents)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                apInvoice.GetByKey(docEntry);
                var costCenterCode = apInvoice.UserFields.Fields.Item("U_CostCenter").Value.ToString();
                if (string.IsNullOrWhiteSpace(costCenterCode))
                {
                    return;
                }

                int paymentEntry;

                Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT DocNum FROM VPM2 where DocEntry = {docEntry} and InvType = 18"));
                if (recSet.EoF)
                {
                    return;
                }
                else
                {
                    paymentEntry = int.Parse(recSet.Fields.Item("DocNum").Value.ToString());
                }


                Payments outgoingPayment = (SAPbobsCOM.Payments)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                outgoingPayment.GetByKey(paymentEntry);

                if (outgoingPayment.DocType != BoRcptTypes.rSupplier)
                {
                    return;
                }

                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT TransId FROM OVPM WHERE DocEntry = {paymentEntry}"));

                int jdtTransId = int.Parse(DiManager.Recordset.Fields.Item("TransId").Value.ToString());

                JournalEntries journalEntry = (SAPbobsCOM.JournalEntries)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                journalEntry.GetByKey(jdtTransId);

                outgoingPayment.Invoices.SetCurrentLine(0);


                if (string.IsNullOrWhiteSpace(costCenterCode))
                {
                    return;
                }

                for (int i = 0; i < journalEntry.Lines.Count; i++)
                {
                    journalEntry.Lines.SetCurrentLine(i);
                    if (journalEntry.Lines.AccountCode == "1430")
                    {
                        switch (DiManager.EmployeeDimension)
                        {
                            case DiManager.Dimension.Dimention1 :
                                journalEntry.Lines.CostingCode = costCenterCode;
                                journalEntry.Update();
                                break;
                            case DiManager.Dimension.Dimention2 :
                                journalEntry.Lines.CostingCode2 = costCenterCode;
                                journalEntry.Update();
                                break;
                            case DiManager.Dimension.Dimention3 :
                                journalEntry.Lines.CostingCode3 = costCenterCode;
                                journalEntry.Update();
                                break;
                            case DiManager.Dimension.Dimention4 :
                                journalEntry.Lines.CostingCode4 = costCenterCode;
                                journalEntry.Update();
                                break;
                            case DiManager.Dimension.Dimention5 :
                                journalEntry.Lines.CostingCode5 = costCenterCode;
                                journalEntry.Update();
                                break;
                        }
                    }
                }
            }
        }
    }
}
