using System;
using System.Collections.Generic;
using System.Xml;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace CostCenterOutgoing
{
    [FormAttribute("CostCenterOutgoing.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Payments outgoingPayment = (SAPbobsCOM.Payments)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            outgoingPayment.GetByKey(9);

            DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT TransId FROM OVPM WHERE DocEntry = {9}"));

            int jdtTransId = int.Parse(DiManager.Recordset.Fields.Item("TransId").Value.ToString());

            JournalEntries journalEntry = (SAPbobsCOM.JournalEntries)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            journalEntry.GetByKey(jdtTransId);
           var x1 = journalEntry.GetAsXML();
            var x = outgoingPayment.Invoices.DistributionRule2;
        
        }
    }
}