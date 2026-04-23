using System;
using System.IO;
using System.Reflection;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library.EntityDefinition;

namespace Thermo.SampleManager.ObjectModel
{
	/// <summary>
	/// Defines extended business logic and manages access to the Result entity.
	/// </summary> 
	[SampleManagerEntity(EntityName)]
	public class ResultExtension : Thermo.SampleManager.ObjectModel.Result
	{
		#region Member Variables

		private IEntityCollection m_AppliedLimits;
	
		#endregion

		#region Properties

		/// <summary>
		/// Gets the applied limits.
		/// </summary>
		/// <value>The applied limits.</value>
		[PromptCollection(MlpValuesBase.EntityName, false)]
		public IEntityCollection AppliedLimits
		{
			get
			{
				if (m_AppliedLimits != null) return m_AppliedLimits;

				// Get to the limits via the Test > Sample > MlpComponent > MlpValue

				string mlp = TestNumber.Sample.Product;
				if (string.IsNullOrEmpty(mlp)) return null;

				// Product is assigned, get the MLP Component for this result

				IQuery query = EntityManager.CreateQuery(MlpComponentsBase.EntityName);

				query.AddEquals(MlpComponentsPropertyNames.ProductId, mlp);
				query.AddEquals(MlpComponentsPropertyNames.ProductVersion, TestNumber.Sample.ProductVersion);
				query.AddEquals(MlpComponentsPropertyNames.AnalysisId, TestNumber.Analysis.Identity);
				query.AddEquals(MlpComponentsPropertyNames.ComponentName, ResultName);

				IEntityCollection mlpComponents = EntityManager.Select(MlpComponentsBase.EntityName, query);
				if (mlpComponents.Count == 0) return null;

				// There are limits...

				MlpComponentsBase mlpComponent = (MlpComponentsBase) mlpComponents[0];

				// Get the limits from MLP_VALUES

				query = EntityManager.CreateQuery(MlpValuesBase.EntityName);
				query.AddEquals(MlpValuesPropertyNames.EntryCode, mlpComponent.EntryCode);

				m_AppliedLimits = EntityManager.Select(MlpValuesBase.EntityName, query);

				return m_AppliedLimits;
			}
		}

		#endregion

		}
}