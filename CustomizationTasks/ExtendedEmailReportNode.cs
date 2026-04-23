using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Thermo.SampleManager.Library.EntityDefinition;
using Thermo.SampleManager.ObjectModel;
using Thermo.SampleManager.Server.Workflow.Nodes;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Library.ClientControls;
using Thermo.SampleManager.Server.Workflow.Attributes;
using Thermo.SampleManager.Server.Workflow.Definition;
using Thermo.SampleManager.Server;
using DevExpress.Data.Svg;
using Thermo.SampleManager.Server.FormulaFunctionService;
using Thermo.SampleManager.Server.Workflow;
using DevExpress.Data.Async;
using static System.Net.WebRequestMethods;
using System.Linq;
using Thermo.SampleManager.ObjectModel.Sources;
using System.Runtime.CompilerServices;
using Thermo.SampleManager.Library.ClientControls.Browse;
using System.Net.Mail;
using Thermo.Framework.Core;
using Thermo.SampleManager.Core.Exceptions;

namespace Customization.Tasks
{

    /// <summary>
    /// Workflow node to Export to cloud Storage
    /// </summary>
    [FollowsTag(WorkflowNodeTypeInternal.TagOutput)]
    [FollowsTag(WorkflowNodeTypeInternal.TagDataOutput)]
    [Tag(WorkflowNodeTypeInternal.TagDestination)]
    [WorkflowNode(NodeTypeNew,
        NodeName,
        NodeCategory,
        NodeIcon,
        NodeDescription,
        NodeMessageGroup)]
    public class FileNameChangeMailNode : MailNode
    {
        /// <summary>
        /// The node type
        /// </summary>
        public const string NodeTypeNew = "MAIL_DESTINATION_FILE";

        /// <summary>
        /// The node name
        /// </summary>
        public const string NodeName = "NodeRenameFileinEmailNodeName";

        /// <summary>
        /// The node category
        /// </summary>
        public const string NodeCategory = "NodeRenameFileinEmailNodeCategory";

        /// <summary>
        /// The node icon
        /// </summary>
        public const string NodeIcon = "INT_MAIL_NEW";

        /// <summary>
        /// The node description
        /// </summary>
        public const string NodeDescription = "NodeRenameFileinEmailNodeDescription";

        /// <summary>
        /// The node message group
        /// </summary>
        public const string NodeMessageGroup = "ArcadiumMessages";


        /// <summary>
        /// The From Line parameter entity
        /// </summary>
        public const string ParameterFromLine = "FROMLINE";



        /// <summary>
        /// The FileName parameter entity
        /// </summary>
        public const string ParameterFileName = "FILENAME";

        /// <summary>
        /// The From Line parameter entity
        /// </summary>
        public const string ParameterCustomerLine = "CUSTOMERLINE";



        /// <summary>
        /// Mail override node<see cref="ImportFromCloudStorageNode"/> class.
        /// </summary>
        /// <param name="node">The node.</param>
        /// 


        public FileNameChangeMailNode(WorkflowNodeInternal node) : base(node)
        {
        }



        /// <summary>
        /// Gets or sets the Bucket.
        /// </summary>
        /// <value>
        /// The Folder.
        /// </value>
        [FormulaParameter(
            FileNameChangeMailNode.ParameterFromLine,
            "NodeRenameFileinEmailNodeParameterFromLine",
            "NodeRenameFileinEmailNodeParameterFromLineDescription",
            MessageGroup = "ArcadiumMessages", Mandatory = false)]
        [PromptText]
        public string FROMLINE
        {
            get
            {
                return base.GetParameterBagValue<string>(FileNameChangeMailNode.ParameterFromLine);
            }

            set
            {
                base.SetParameterBagValue(FileNameChangeMailNode.ParameterFromLine, value);
            }
        }


        /// <summary>
        /// Gets or sets the File.
        /// </summary>
        /// <value>
        /// The Folder.
        /// </value>
        [FormulaParameter(
            FileNameChangeMailNode.ParameterFileName,
            "NodeRenameFileinEmailNodeParameterFilename",
            "NodeRenameFileinEmailNodeParameterFilenameDescription",
            MessageGroup = "ArcadiumMessages", Mandatory = false)]
        [PromptText]
        public string FILENAME
        {
            get
            {
                return base.GetParameterBagValue<string>(FileNameChangeMailNode.ParameterFileName);
            }

            set
            {
                base.SetParameterBagValue(FileNameChangeMailNode.ParameterFileName, value);
            }
        }

        /// <summary>
        /// Gets or sets the Customer.
        /// </summary>
        /// <value>
        /// The Folder.
        /// </value>
        [FormulaParameter(
            FileNameChangeMailNode.ParameterCustomerLine,
            "NodeRenameFileinEmailNodeParameterCustomer",
            "NodeRenameFileinEmailNodeParameterCustomerDescription",
            MessageGroup = "ArcadiumMessages", Mandatory = false)]
        [PromptText]
        public string CUSTOMER
        {
            get
            {
                return base.GetParameterBagValue<string>(FileNameChangeMailNode.ParameterCustomerLine);
            }

            set
            {
                base.SetParameterBagValue(FileNameChangeMailNode.ParameterCustomerLine, value);
            }
        }



        /// <summary>
        /// The main execution function of the node.
        /// </summary>
        /// <returns>
        /// The result if the operation was valid.
        /// </returns>
        public override bool PerformNode()
        {


            try
            {
                TracePerformNode();
                if (!base.Properties.ContainsKey("$OutputFile"))
                {
                    TraceDebug("TraceMailNodeNoOutput");
                    return base.PerformNode();
                }
                string text = (string)base.Properties["$OutputFile"];
                if (text == null || !System.IO.File.Exists(text))
                {
                    TraceDebug("TraceMailNodeNoOutput");
                    return false;
                }
                int count = base.Properties.Errors.Count;
                MailMessage mailMessage = BuildMailMessage();
                if (mailMessage == null)
                {
                    return false;
                }
                if (base.Properties.Errors.Count > 0 && base.Properties.Errors.Count != count)
                {
                    base.Properties.Errors.RemoveAt(base.Properties.Errors.Count - 1);
                }
                TraceDebug("TraceMailNodeAttachment", text);
                string exten = Path.GetExtension(text);

                using System.Net.Mail.Attachment item = new System.Net.Mail.Attachment(text);


                item.Name = GetFormulaText(this.FILENAME) + exten;

                mailMessage.Attachments.Add(item);
                string fromadd = GetFormulaText(FROMLINE);
                MailAddressCollection addressesfrom = new MailAddressCollection();
                MailAddressCollection addressescustomer = new MailAddressCollection();
                addressesfrom.Add(fromadd);
                string customer = GetFormulaText(CUSTOMER);
                string custemails = string.Empty;
                if ((customer != null) && !(string.IsNullOrWhiteSpace(customer)))
                {
                    Customer cus = (Customer)EntityManager.Select(Customer.EntityName, customer);
                    if (cus != null && cus.IsValid())
                    {
                        foreach (Contact contact in cus.Contacts)
                        {
                            if (!string.IsNullOrEmpty(contact.Email))
                            {
                                addressescustomer.Add(contact.Email);
                                if (string.IsNullOrEmpty(custemails))
                                    custemails = contact.Email;
                                else
                                {
                                    custemails += "," + contact.Email;
                                }

                            }
                        }
                    }
                }



                mailMessage.From = addressesfrom[0];
                mailMessage.Sender = addressesfrom[0];
                mailMessage.To.Insert(0, addressesfrom[0]);
                if (addressescustomer.Count > 0)
                {
                    mailMessage.To.Add(custemails);
                }
                TraceInfo("TraceMailNodeSend", mailMessage.To);
                base.Library.Utils.Mail(mailMessage);
                TraceDebug("TraceMailNodeSend", mailMessage.To);
            }
            catch (Exception ex)
            {
                TraceError("TraceMailNodeError", ex.Message);
            }
            return true;



        }

        public new static void GetPromptDistributionPrinter(SampleManagerTask task, WorkflowNodeInternal node, IEntity rowEntity, DataGridColumn gridProperty)
        {
            IQuery query = task.EntityManager.CreateQuery("PRINTER");
            query.AddEquals("DeviceType", "MAIL");
            query.AddOr();
            query.AddEquals("DeviceType", "MAIL_OPER");
            query.AddOr();
            query.AddEquals("DeviceType", "MAIL_CONT");
            query.AddOr();
            query.AddEquals("DeviceType", "DIST");
            EntityBrowse entityBrowse = task.BrowseFactory.CreateEntityBrowse(query);
            gridProperty.SetCellBrowse(rowEntity, entityBrowse);
        }


        public override string AutoName()
        {
            return this.Library.Message.GetMessage("ArcadiumMessages", "NodeExportToCloudName");
        }
    }
}
