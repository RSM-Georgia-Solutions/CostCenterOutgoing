using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace CostCenterOutgoing
{
    [FormAttribute("426", "OutgoingPayment.b1f")]
    class OutgoingPayment : SystemFormBase
    {
        public OutgoingPayment()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            if (pVal.ActionSuccess)
            {
                int docEntry = 0;
                try
                {
                    string xmlObjectKey = pVal.ObjectKey;
                    XElement xmlnew = XElement.Parse(xmlObjectKey);
                    XElement xElement = xmlnew.Element("DocEntry");
                    if (xElement != null)
                    {
                        docEntry = int.Parse(xElement.Value);
                        Payments outgoingPayment = (SAPbobsCOM.Payments)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        outgoingPayment.GetByKey(docEntry);
                      
                        if (outgoingPayment.DocType != BoRcptTypes.rSupplier)
                        {
                            return;
                        }

                        DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT TransId FROM OVPM WHERE DocEntry = {docEntry}"));

                        int jdtTransId = int.Parse(DiManager.Recordset.Fields.Item("TransId").Value.ToString());

                        JournalEntries journalEntry = (SAPbobsCOM.JournalEntries)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        journalEntry.GetByKey(jdtTransId);

                        outgoingPayment.Invoices.SetCurrentLine(0);
                        var costCenter = outgoingPayment.Invoices.DistributionRule2;
                        switch (DiManager.EmployeeDimension)
                        {
                            case DiManager.Dimension.Dimention1:
                                costCenter = outgoingPayment.Invoices.DistributionRule;
                                break;
                            case DiManager.Dimension.Dimention2:
                                costCenter = outgoingPayment.Invoices.DistributionRule2;
                                break;
                            case DiManager.Dimension.Dimention3:
                                costCenter = outgoingPayment.Invoices.DistributionRule3;
                                break;
                            case DiManager.Dimension.Dimention4:
                                costCenter = outgoingPayment.Invoices.DistributionRule4;
                                break;
                            case DiManager.Dimension.Dimention5:
                                costCenter = outgoingPayment.Invoices.DistributionRule5;
                                break;
                        }

                        if (string.IsNullOrWhiteSpace(costCenter))
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
                                        journalEntry.Lines.CostingCode = costCenter;
                                        journalEntry.Update();
                                        break;
                                    case DiManager.Dimension.Dimention2 :
                                        journalEntry.Lines.CostingCode2 = costCenter;
                                        journalEntry.Update();
                                        break;
                                    case DiManager.Dimension.Dimention3 :
                                        journalEntry.Lines.CostingCode3 = costCenter;
                                        journalEntry.Update();
                                        break;
                                    case DiManager.Dimension.Dimention4 :
                                        journalEntry.Lines.CostingCode4 = costCenter;
                                        journalEntry.Update();
                                        break;
                                    case DiManager.Dimension.Dimention5 :
                                        journalEntry.Lines.CostingCode5 = costCenter;
                                        journalEntry.Update();
                                        break;
                                    default :
                                        throw new ArgumentOutOfRangeException();
                                }
                            }
                        }
                    }
                    else
                    {
                        Application.SBO_Application.SetStatusBarMessage("Outgoing Payment DocEntry Not Found",
                            BoMessageTime.bmt_Short);
                    }

                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                }
            }

        }

        private void OnCustomInitialize()
        {

        }

        private Button Button0;

        private void Button0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
    }
}
