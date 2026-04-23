using System;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Library.EntityDefinition;
using Thermo.SampleManager.ObjectModel;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Common.Workflow;

namespace Customization.Tasks
{
    [SampleManagerTask("JobCreation", "WorkflowCallback")]
    public class JobCreation : SampleManagerTask
    {
        protected override void SetupTask()
        {
            base.SetupTask();
            var sample = GetSampleFromContext();
            Library.Utils.FlashMessage("DBG JobCreation → Start", "JobCreation");
            if (!BaseEntity.IsValid(sample)) { Exit(null); return; }

            var delivery = GetDeliveryNumber(sample);
            var lineItemForDbg = GetLineItemNumber(sample);
            var productIdForDbg = sample.Product != null ? sample.Product.ToString() : string.Empty;
            var sampleIdForDbg = string.Empty; try { sampleIdForDbg = ((IEntity)sample).GetString("ID_NUMERIC"); } catch { }
            Library.Utils.FlashMessage($"DBG Sample={sampleIdForDbg}; Product={productIdForDbg}; Delivery='{delivery}'; LineItem={lineItemForDbg}", "JobCreation");
            if (string.IsNullOrWhiteSpace(delivery)) { Exit(null); return; }

            if (sample.JobName != null && sample.JobName.IsValid()) { var name = sample.JobName.ToString(); TrySetWorkflowValue("JOB_NAME", name); Library.Utils.FlashMessage($"JOB_NAME: {name}", "JobCreation"); Exit(name); return; }

            var job = FindMatchingJob(sample, delivery);
            if (BaseEntity.IsValid(job)) { TrySetWorkflowValue("JOB_NAME", job.JobName); Library.Utils.FlashMessage($"JOB_NAME: {job.JobName}", "JobCreation"); Exit(job.JobName); return; }
            TrySetWorkflowValue("JOB_NAME", null);
            Library.Utils.FlashMessage("JOB_NAME: (null)", "JobCreation");
            Exit(null);
        }

        private void TrySetWorkflowValue(string key, object value)
        {
            var bag = GetOrCreateWorkflowBag();
            if (bag == null) return;
            try
            {
                var bagType = bag.GetType().FullName;
                Library.Utils.FlashMessage($"DBG JobCreation → Bag type: {bagType}", "JobCreation");
                bag.SetGlobalVariable("JOB_NAME", value);
                bag["sample.JOB_NAME"] = value;
                bag["samples.0.JOB_NAME"] = value;
                var countVal = value == null ? 0 : 1;
                bag["samples.Count"] = countVal;
                var v = value == null ? "(null)" : value.ToString();
                Library.Utils.FlashMessage($"DBG JobCreation → Set Global.JOB_NAME={v}; sample.JOB_NAME={v}; samples.0.JOB_NAME={v}; samples.Count={countVal}", "JobCreation");
            }
            catch { }
        }

        private IWorkflowPropertyBag GetOrCreateWorkflowBag()
        {
            try
            {
                var ctx = this.Context;
                if (ctx == null) return null;
                var ctxType = ctx.GetType();
                // 1) Propiedades que ya conocemos por nombre
                var p1 = ctxType.GetProperty("WorkflowPropertyBag");
                if (p1 != null)
                {
                    var v = p1.GetValue(ctx, null);
                    if (v is IWorkflowPropertyBag wb1) return wb1;
                }
                var p2 = ctxType.GetProperty("PropertyBag");
                if (p2 != null)
                {
                    var v2 = p2.GetValue(ctx, null);
                    if (v2 is IWorkflowPropertyBag wb2) return wb2;
                }
                // 2) Cualquier propiedad pública que implemente IWorkflowPropertyBag
                foreach (var pi in ctxType.GetProperties())
                {
                    try
                    {
                        var val = pi.GetValue(ctx, null);
                        if (val is IWorkflowPropertyBag wbAny) return wbAny;
                    }
                    catch { }
                }
                // 3) Campos públicos que implementen IWorkflowPropertyBag
                foreach (var fi in ctxType.GetFields())
                {
                    try
                    {
                        var val = fi.GetValue(ctx);
                        if (val is IWorkflowPropertyBag wbField) return wbField;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private int GetLineItemNumber(Sample sample)
        {
            try
            {
                int n = 0; int.TryParse(((IEntity)sample).GetString("TF_LINE_ITEM_NUMBER"), out n); return n;
            }
            catch { return 0; }
        }

        private Sample GetSampleFromContext()
        {
            if (Context.SelectedItems != null && Context.SelectedItems.Count > 0 && Context.SelectedItems[0] is Sample)
                return (Sample)Context.SelectedItems[0];
            return null;
        }

        private string GetDeliveryNumber(Sample sample)
        {
            try { return ((IEntity)sample).GetString("TF_DELIVERY_NUMBER"); } catch { return string.Empty; }
        }

        private JobHeader FindMatchingJob(Sample sample, string delivery)
        {
            string productIdentityForJob = sample.Product != null ? sample.Product.ToString() : string.Empty;
            if (string.IsNullOrEmpty(productIdentityForJob)) return null;
            int deliveryNumber = 0; int.TryParse(delivery, out deliveryNumber);
            int lineItemNumber = 0; try { int.TryParse(((IEntity)sample).GetString("TF_LINE_ITEM_NUMBER"), out lineItemNumber); } catch { lineItemNumber = 0; }
            var q = EntityManager.CreateQuery(JobHeaderBase.EntityName);
            q.AddEquals("TF_PRODUCT", productIdentityForJob);
            q.AddAnd();
            q.AddEquals("TF_DELIVERY_NUMBER", deliveryNumber);
            q.AddAnd();
            q.AddEquals("TF_LINE_ITEM_NUMBER", lineItemNumber);
            var jobs = EntityManager.Select(JobHeaderBase.EntityName, q);
            if (jobs != null && jobs.Count > 0)
            {
                return jobs.GetFirst() as JobHeader;
            }
            return null;
        }


    }
}


