using System;
using System.IO;
using Thermo.SampleManager.Common;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Library.EntityDefinition;
using Thermo.SampleManager.ObjectModel;
using Thermo.SampleManager.Server.Workflow.Attributes;
using Thermo.SampleManager.Server.Workflow.Definition;
using Thermo.SampleManager.Server.Workflow.Nodes;

namespace Arcadium.Workflow.Nodes
{
    [WorkflowNode(
        "ARCADIUM_REPORT_GENERATOR",
        "GenerateReportNodeName",
        "ArcadiumCategory",
        "REPORT_GENERATOR",
        "GenerateReportNodeDescription",
        MessageGroup)]
    [FollowsTag(WorkflowNodeTypeInternal.TagEntryPoint)]
    [FollowsTag(WorkflowNodeTypeInternal.TagRepeat)]
    [FollowsTag(WorkflowNodeTypeInternal.TagData)]
    public class ArcadiumGenerateReportNode : Node
    {
        private const string MessageGroup = "ArcadiumPrintingMessages";

        public ArcadiumGenerateReportNode(WorkflowNodeInternal node) : base(node) { }

        [FormulaParameter("DELIVERY_NUMBER",
                          "Delivery Number",
                          "Enter the delivery number",
                          Mandatory = true,
                          MessageGroup = MessageGroup)]
        [PromptText]
        public string DeliveryNumber
        {
            get => GetParameterBagValue<string>("DELIVERY_NUMBER");
            set => SetParameterValue("DELIVERY_NUMBER", value);
        }

        public override bool PerformNode()
        {
            string rawDelivery = DeliveryNumber;
            string rawFileLoc = "E:\\Thermo\\SampleManager\\Server\\VGSM\\COA";
            string deliveryId = GetFormulaText(rawDelivery);
            string[] props = deliveryId.Split(",");
            string fileLocation = GetFormulaText(rawFileLoc);

            try
            {
                IQuery q = EntityManager.CreateQuery("SAMP_TEST_RESULT_WITH_LIMITS");
                q.AddEquals(SampTestResultWithLimitsPropertyNames.DeliveryNumber, props[0]);
                var results = EntityManager.Select(q);

                string deliveryName = $"F{props[2]}-{props[0]}-{props[1]}-V{props[3]}-{DateTime.Now.ToString("yyyyMMdd")}";

                if (results.Count == 0)
                    return false;

                if (!Directory.Exists(fileLocation))
                    Directory.CreateDirectory(fileLocation);

                string reportPath = Path.Combine(fileLocation, deliveryName);
                GenerateReport(results, reportPath);
                string fileName = Path.GetFileName(reportPath);
                Library.Utils.FlashMessage(
                    $"El reporte {fileName} se creo en la direccion: {fileLocation}",
                    "Report Generated",
                    MessageButtons.OK,
                    MessageIcon.Information,
                    MessageDefaultButton.Button1);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private void GenerateReport(IEntityCollection records, string fileLocation)
        {
            var tmpl = EntityManager.SelectLatestVersion<ReportTemplate>("COA_RT");
            if (!BaseEntity.IsValid(tmpl))
                throw new InvalidOperationException("Template 'COA_RT' not found");
            Library.Reporting.PrintReportToServerFile(
                tmpl,
                records,
                new ReportOptions(),
                PrintFileType.Pdf,
                fileLocation);
        }

        public override string AutoName() =>
            FormatMessage("GenerateReportNodeName");

        private string FormatMessage(string key, params object[] args) =>
            Library.Message.GetMessage(MessageGroup, key, args);
    }
}
