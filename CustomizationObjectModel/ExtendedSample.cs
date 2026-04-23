using System.Collections.Generic;
using System.Linq;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Library.EntityDefinition;

namespace Thermo.SampleManager.ObjectModel
{
	/// <summary>
	/// Defines extended business logic and manages access to the HAZARD entity.
	/// </summary> 
	[SampleManagerEntity( "SAMPLE" )]
	public class ExtendedSample : Sample
	{
				
		/// <summary>
		/// Gets lot samples results.
		/// </summary>
		/// <value>
		/// The samples results.
		/// </value>
		[PromptCollection(TLotsampTestRslBase.EntityName, false, StopAutoPublish = true)]
		public IEntityCollection LotSampleResults
		{
			get
			{
				var query = EntityManager.CreateQuery(TLotsampTestRslBase.EntityName);
				query.AddEquals(TLotsampTestRslPropertyNames.IdNumeric, IdNumeric);
                query.AddNotEquals(TLotsampTestRslPropertyNames.TestStatus, "X");
                query.AddOrder(TLotsampTestRslPropertyNames.TestOrder, true);
                query.AddOrder(TLotsampTestRslPropertyNames.ComponentOrder, true);
                return EntityManager.Select(query);
			}
		}


        /// <summary>
        /// Gets Components that failed the Generic MLP level.
        /// </summary>
        /// <value>
        /// Component list.
        /// </value>
        [PromptText]
        public string GenericFailedComp
        {
            get
            {

                string cl = string.Empty;                

                IQuery qry = EntityManager.CreateQuery(TLotsampTestRslBase.EntityName);
                qry.AddEquals(TLotsampTestRslPropertyNames.IdNumeric, IdNumeric);                
                qry.AddEquals(TLotsampTestRslPropertyNames.GenericOos, true);
                qry.AddEquals(TLotsampTestRslPropertyNames.ReportFlag, "REP");
                qry.AddNotEquals(TLotsampTestRslPropertyNames.TestStatus, "X");
                qry.AddOrder(TLotsampTestRslPropertyNames.TestOrder, true);
                qry.AddOrder(TLotsampTestRslPropertyNames.ComponentOrder, true);
                IEntityCollection CompList = EntityManager.Select(qry);

                if (CompList.ActiveItems.Count > 0)
                {
                    foreach (TLotsampTestRslBase comp in CompList)
                    {
                        if (string.IsNullOrEmpty(cl)) cl = comp.ComponentName.Trim();
                        else cl += "; " + comp.ComponentName.Trim();
                    }                    
                }

                return cl;
                                
            }
        }

        /// <summary>
		/// Gets sample MLP levels.
		/// </summary>
		/// <value>
		/// All level status.
		/// </value>
		[PromptCollection(TSampMlpLevelBase.EntityName, false, StopAutoPublish = true)]
        public IEntityCollection LotSampLevel
        {
            get
            {
                var query = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                query.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);                
                return EntityManager.Select(query);
            }           

        }

        /// <summary>
		/// Determine if ALL sample's levels are out of spec.
		/// </summary>
		/// <value>
		///   <c>true</c> if all levels are OOS; otherwise, <c>false</c>.
		/// </value>
		[PromptBoolean]
        public bool OosAllLevels
        {
            get
            {
                var qry = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                qry.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);
                int totalRows = EntityManager.SelectCount(qry);

                var qryOos = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                qryOos.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);
                qryOos.AddEquals(TSampMlpLevelPropertyNames.OutOfSpec, true);
                int oosRows = EntityManager.SelectCount(qryOos);

                if (totalRows == oosRows)
                    return true;
                else
                    return false;

            }
        }

        /// <summary>
        /// Determine if ALL sample's levels are in spec.
        /// </summary>
        /// <value>
        ///   <c>true</c> if all levels are in Spec; otherwise, <c>false</c>.
        /// </value>
        [PromptBoolean]
        public bool OosNoneLevels
        {
            get
            {
                IQuery qry = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                qry.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);
                int totalRows = EntityManager.SelectCount(qry);

                IQuery qryInSpec = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                qryInSpec.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);
                qryInSpec.AddEquals(TSampMlpLevelPropertyNames.OutOfSpec, false);
                int inRows = EntityManager.SelectCount(qryInSpec);

                if (totalRows == inRows)
                    return true;
                else
                    return false;

            }

        }

        /// <summary>
        /// Read sample's Generic level spec.
        /// </summary>
        /// <value>
        ///   <c>true</c> if Generic level is out of Spec; otherwise, <c>false</c>.
        /// </value>
        [PromptBoolean]
        public bool OosGenericLevel
        {
            get
            {
                string genlvl = Library.Environment.GetGlobalString("T_FENIX_GENERIC_MLP_LEVEL").Trim();
                
                IQuery qry = EntityManager.CreateQuery(TSampMlpLevelBase.EntityName);
                qry.AddEquals(TSampMlpLevelPropertyNames.IdNumeric, IdNumeric);
                qry.AddEquals(TSampMlpLevelPropertyNames.LevelId, genlvl);
                TSampMlpLevelBase genLevel = EntityManager.Select(qry).GetFirst() as TSampMlpLevelBase;
                
                if (BaseEntity.IsValid(genLevel)) return genLevel.OutOfSpec;
                else return false;

            }

        }


    }
}
