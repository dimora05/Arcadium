/*
    PivotResultsElasticTask
    Autor: Danny Loria
    Fecha: 2025-10-27
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Thermo.SampleManager.Common;
using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library;
using Thermo.SampleManager.Library.ClientControls;
using Thermo.SampleManager.Library.EntityDefinition;
using Thermo.SampleManager.Library.FormDefinition;
using Thermo.SampleManager.ObjectModel;
using Thermo.SampleManager.Tasks;
using Thermo.SampleManager.Server;



namespace Customization.Tasks
{
    [SampleManagerTask("PivotResultsElasticTask")]
    public class PivotResultsElasticTask : DefaultFormTask
    {
        private FormpivotResultsElastic form;
        private const string AdvancedRoleId = "RESULT_VIEW_LAB";


        protected override void MainFormCreated()
        {
            base.MainFormCreated();

            form = (FormpivotResultsElastic)MainForm;

            form.btnSearch.ClickAndWait += BtnBuscar_Click;
            form.btnClearFilters.ClickAndWait += BtnClearFilters_Click;
            form.btnClearResults.ClickAndWait += BtnClearResults_Click;
            form.btnExport.ClickAndWait += BtnExport_Click;

            form.pebPlanta.EntityChanged += PebPlanta_EntityChanged;
            form.pebProceso.EntityChanged += PebProceso_EntityChanged;
            form.pebEtapa.EntityChanged += PebEtapa_EntityChanged;
            form.pebPuntoMuestreo.EntityChanged += PebPunto_ValueChanged;

            InitializeDefaults();
            form.pebPlanta.Value = null; form.pebPlanta.Entity = null;
            form.pebProceso.Value = null; form.pebProceso.Entity = null;
            form.pebEtapa.Value = null; form.pebEtapa.Entity = null;
            form.pebPuntoMuestreo.Value = null; form.pebPuntoMuestreo.Entity = null;

            // Filtro inicial para Plantas
            IQuery qPlantas = EntityManager.CreateQuery("LOCATION");
            qPlantas.AddEquals("LocationType", "PLANTA");
            form.pebPlanta.Browse = BrowseFactory.CreateEntityBrowse(qPlantas);

            SetHierarchyEnabled(null, null, null);
            ApplyAdvancedVisibility();
        }

        private bool IsWebClient
        {
            get
            {
                try
                {
                    return Library != null
                        && Library.Environment != null
                        && Library.Environment.GetGlobalInt("CLIENT_TYPE") == 1;
                }
                catch
                {
                    return false;
                }
            }
        }

        private void ShowInfo(string message, string title)
        {
            if (IsWebClient)
            {
                Library.Utils.ShowAlert(title, message, NotificationType.General);
            }
            else
            {
                Library.Utils.FlashMessage(message, title);
            }
        }

        private void ShowError(string message, string title)
        {
            if (IsWebClient)
            {
                Library.Utils.ShowAlert(title, message, NotificationType.Error);
            }
            else
            {
                Library.Utils.FlashMessage(message, title);
            }
        }

        private void ApplyAdvancedVisibility()
        {
            bool canSeeAdvanced = false;
            var currentUser = Library != null && Library.Environment != null ? Library.Environment.CurrentUser as Personnel : null;
            if (currentUser != null)
            {
                var role = EntityManager.Select(RoleHeader.EntityName, AdvancedRoleId) as RoleHeaderBase;
                if (role != null)
                {
                    try { canSeeAdvanced = currentUser.HasRole(role); } catch { canSeeAdvanced = false; }
                }
            }

            var gb = form.gbAvanzado as VisualControl;
            if (gb != null) gb.Visible = canSeeAdvanced;

            var chk = form.chkCompletadas as VisualControl;
            if (chk != null) chk.Visible = canSeeAdvanced;
        }

        private void GetDateRange(out DateTime start, out DateTime end)
        {
            start = DateTime.MinValue;
            end = DateTime.MaxValue;
            var sdt = form.deFechaInicio.Date;
            if (!sdt.Equals(Thermo.Framework.Core.NullableDateTime.CreateNullInstance()))
            {
                start = new DateTime(sdt.Year, sdt.Month, sdt.Day, 0, 0, 0);
            }
            var edt = form.deFechaTermino.Date;
            if (!edt.Equals(Thermo.Framework.Core.NullableDateTime.CreateNullInstance()))
            {
                end = new DateTime(edt.Year, edt.Month, edt.Day, 23, 59, 59);
            }
            if (start > end)
            {
                var tmp = start; start = end; end = tmp;
            }
        }


        private void BtnClearFilters_Click(object sender, EventArgs e)
        {
            form.pebPuntoMuestreo.Entity = null; form.pebPuntoMuestreo.Value = null;
            form.pebEtapa.Entity = null; form.pebEtapa.Value = null;
            form.pebProceso.Entity = null; form.pebProceso.Value = null;
            form.pebPlanta.Entity = null; form.pebPlanta.Value = null;
            form.chkCompletadas.Checked = false;
            SetHierarchyEnabled(null, null, null);
        }

        private void BtnClearResults_Click(object sender, EventArgs e)
        {
            var grid = form.ugdPivotResults;
            grid.ClearRows();
            grid.ClearColumns();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (form.ugdPivotResults == null || form.ugdPivotResults.Rows == null || form.ugdPivotResults.Rows.Count == 0)
                {
                    ShowInfo("No hay filas para exportar.", "Exportar");
                    return;
                }

                string defaultName = $"pivot-results-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
                var clientFilePath = Library.Utils.PromptForFile("Guardar CSV", "CSV Files|*.csv|All Files|*.*", true, defaultName);
                if (string.IsNullOrWhiteSpace(clientFilePath))
                {
                    return;
                }

                var sb = new StringBuilder();
                Func<string, string> esc = s =>
                {
                    if (s == null) return "";
                    var t = s.Replace("\"", "\"\"");
                    return "\"" + t + "\"";
                };

                for (int c = 0; c < form.ugdPivotResults.Columns.Count; c++)
                {
                    var col = form.ugdPivotResults.Columns[c];
                    sb.Append(esc(col.Caption));
                    if (c < form.ugdPivotResults.Columns.Count - 1) sb.Append(",");
                }
                sb.AppendLine();

                foreach (UnboundGridRow row in form.ugdPivotResults.Rows)
                {
                    for (int c = 0; c < form.ugdPivotResults.Columns.Count; c++)
                    {
                        var col = form.ugdPivotResults.Columns[c];
                        var v = row.GetValue(col.Name);
                        var val = v == null ? string.Empty : v.ToString();
                        sb.Append(esc(val));
                        if (c < form.ugdPivotResults.Columns.Count - 1) sb.Append(",");
                    }
                    sb.AppendLine();
                }

                var serverFile = Library.File.GetWriteFile("smp$textreports", defaultName);
                try
                {
                    using (var sw = new StreamWriter(serverFile.FullName, false, Encoding.UTF8))
                    {
                        sw.Write(sb.ToString());
                    }
                }
                catch (Exception ex)
                {
                    ShowError($"No se pudo generar el archivo: {ex.Message}", "Exportar");
                    return;
                }

                try
                {
                    Library.File.TransferToClient(serverFile.FullName, clientFilePath);
                }
                catch (Exception ex)
                {
                    ShowError($"No se pudo copiar al cliente: {ex.Message}", "Exportar");
                    return;
                }

                // En WebClient el popup puede quedar detrás del navegador, así que solo mostramos
                // el mensaje en el cliente rico. En web, se descarga sin mostrar diálogo modal.
                if (!IsWebClient)
                {
                    ShowInfo($"Archivo guardado en:\n{clientFilePath}", "Exportar");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error al exportar: {ex.Message}", "Exportar");
            }
        }

        private void PebPlanta_EntityChanged(object sender, EntityChangedEventArgs e)
        {
            form.pebProceso.Entity = null; form.pebProceso.Value = null;
            form.pebEtapa.Entity = null; form.pebEtapa.Value = null;
            form.pebPuntoMuestreo.Entity = null; form.pebPuntoMuestreo.Value = null;

            var planta = (e != null ? e.Entity : null) ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
            if (planta == null)
            {
                SetHierarchyEnabled(null, null, null);
                return;
            }

            SetHierarchyEnabled(planta, null, null);
        }

        private void PebProceso_EntityChanged(object sender, EntityChangedEventArgs e)
        {
            form.pebEtapa.Entity = null; form.pebEtapa.Value = null;
            form.pebPuntoMuestreo.Entity = null; form.pebPuntoMuestreo.Value = null;

            var proceso = (e != null ? e.Entity : null) ?? ResolveEntityByValue("LOCATION", form.pebProceso?.Value);
            if (proceso == null)
            {
                var planta = form.pebPlanta?.Entity ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
                SetHierarchyEnabled(planta, null, null);
                return;
            }

            var planta2 = form.pebPlanta?.Entity ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
            SetHierarchyEnabled(planta2, proceso, null);
        }

        private void PebEtapa_EntityChanged(object sender, EntityChangedEventArgs e)
        {
            form.pebPuntoMuestreo.Entity = null; form.pebPuntoMuestreo.Value = null;

            var etapa = (e != null ? e.Entity : null) ?? ResolveEntityByValue("LOCATION", form.pebEtapa?.Value);
            if (etapa == null)
            {
                var planta = form.pebPlanta?.Entity ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
                var proceso = form.pebProceso?.Entity ?? ResolveEntityByValue("LOCATION", form.pebProceso?.Value);
                SetHierarchyEnabled(planta, proceso, null);
                return;
            }

            var planta3 = form.pebPlanta?.Entity ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
            var proceso3 = form.pebProceso?.Entity ?? ResolveEntityByValue("LOCATION", form.pebProceso?.Value);
            SetHierarchyEnabled(planta3, proceso3, etapa);
        }

        private void PebPunto_ValueChanged(object sender, EventArgs e)
        {
            if (form == null || form.pebPuntoMuestreo == null) return;
            if (form.pebPuntoMuestreo.Entity != null) return;
            var ent = ResolveEntityByValue("SAMPLE_POINT", form.pebPuntoMuestreo.Value);
            if (ent != null) form.pebPuntoMuestreo.Entity = ent;
        }

        private void InitializeDefaults()
        {
            var env = Library?.Environment;
            if (env != null && form?.deFechaInicio != null && form?.deFechaTermino != null)
            {
                var today = env.ClientNow.Value.Date;
                var end = new DateTime(today.Year, today.Month, today.Day, 23, 59, 59);
                var startDate = today.AddDays(-7);
                var start = new DateTime(startDate.Year, startDate.Month, startDate.Day, 0, 0, 0);
                form.deFechaInicio.Date = new Thermo.Framework.Core.NullableDateTime(start);
                form.deFechaTermino.Date = new Thermo.Framework.Core.NullableDateTime(end);
            }
        }



        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            var plantaEntForValidation = ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
            if (plantaEntForValidation == null)
            {
                ShowInfo("Seleccione una Planta antes de buscar.", "Buscar");
                return;
            }

            if (form?.deFechaInicio == null || form?.deFechaTermino == null)
            {
                ShowInfo("Seleccione fecha Inicio y Término.", "Buscar");
                return;
            }
            var sdt = form.deFechaInicio.Date;
            var edt = form.deFechaTermino.Date;
            if (sdt.Equals(Thermo.Framework.Core.NullableDateTime.CreateNullInstance()) ||
                edt.Equals(Thermo.Framework.Core.NullableDateTime.CreateNullInstance()))
            {
                ShowInfo("Seleccione fecha Inicio y Término.", "Buscar");
                return;
            }

            form.SetBusy("Por favor espere...", "Buscando");
            string entityName = "SAMP_TEST_RESULT";
            IQuery query = EntityManager.CreateQuery(entityName);

            DateTime start, end;
            GetDateRange(out start, out end);

            string loginDateField = ResolveFieldName(entityName, "LOGIN_DATE", "login_date");

            IEntity plantaEntity = form.pebPlanta?.Entity ?? ResolveEntityByValue("LOCATION", form.pebPlanta?.Value);
            IEntity procesoEntity = form.pebProceso?.Entity ?? ResolveEntityByValue("LOCATION", form.pebProceso?.Value);
            IEntity etapaEntity = form.pebEtapa?.Entity ?? ResolveEntityByValue("LOCATION", form.pebEtapa?.Value);
            IEntity puntoEntity = form.pebPuntoMuestreo?.Entity ?? ResolveEntityByValue("SAMPLE_POINT", form.pebPuntoMuestreo?.Value);
            var samplePoints = BuildSamplingPoints(plantaEntity, procesoEntity, etapaEntity, puntoEntity);

            bool userSelectedLocation = plantaEntity != null || procesoEntity != null || etapaEntity != null || puntoEntity != null;
            if (userSelectedLocation && (samplePoints == null || samplePoints.Count == 0))
            {
                var empty = EntityManager.CreateEntityCollection(entityName);
                BuildPivotGrid(empty);
                SafeClearBusy();
                return;
            }

            string testStatusField = ResolveFieldName(entityName, "TEST_STATUS", "test_status", "STATUS");
            string sampleStatusField = ResolveFieldName(entityName, "STATUS", "SAMPLE_STATUS", "sample_status");
            bool includeCompletedFlag = form?.chkCompletadas != null && form.chkCompletadas.Checked;
            var resultados = BuildResults(entityName, loginDateField, testStatusField, sampleStatusField, start, end, samplePoints, includeCompletedFlag);
            BuildPivotGrid(resultados);
            SafeClearBusy();

        }

        private void BuildPivotGrid(IEntityCollection resultados)
        {
            var grid = form.ugdPivotResults;
            var vcGrid = grid as VisualControl;
            bool beganUpdate = false;
            if (grid != null)
            {
                try { grid.BeginUpdate(); beganUpdate = true; } catch { }
                grid.ClearColumns();
                grid.ClearRows();
            }
            int sampleWidth = EstimateWidth("Muestra");
            int samplePointWidth = EstimateWidth("Punto de Muestreo");
            int sampleStatusWidth = EstimateWidth("Estado de la Muestra");
            int idNumericWidth = EstimateWidth("ID Numérico");
            int loginDateWidth = EstimateWidth("Fecha Login");
            int loteWidth = EstimateWidth("Lote");
            int bagWidth = EstimateWidth("Bolsa");
            int formWidth = EstimateWidth("Formulario");

            UnboundGridColumn colIdNumeric = null, colSample = null, colSamplePoint = null, colLote = null, colBag = null, colForm = null, colLoginDate = null, colSampleStatus = null;
            colIdNumeric = form.ugdPivotResults.AddColumn("IdNumericCol", "ID Numérico", idNumericWidth);
            colSample = form.ugdPivotResults.AddColumn("SampleCol", "Muestra", sampleWidth);
            colSamplePoint = form.ugdPivotResults.AddColumn("SamplePointCol", "Punto de Muestreo", samplePointWidth);
            colLote = form.ugdPivotResults.AddColumn("SampleLoteCol", "Lot", loteWidth);
            colBag = form.ugdPivotResults.AddColumn("SampleBagCol", "Bag", bagWidth);
            colForm = form.ugdPivotResults.AddColumn("SampleFormCol", "Form", formWidth);
            colLoginDate = form.ugdPivotResults.AddColumn("SampleLoginDateCol", "Fecha Login", loginDateWidth);
            colSampleStatus = form.ugdPivotResults.AddColumn("SampleStatusCol", "Estado de la Muestra", sampleStatusWidth);

            if (resultados == null || resultados.Count == 0)
            {
                if (beganUpdate) { try { grid.EndUpdate(); } catch { } }
                SafeClearBusy();
                return;
            }

            string entityName = "T_LOTSAMP_TEST_RSL";
            string sampleIdField = ResolveFieldName(entityName, "ID_TEXT");
            string analysisField = ResolveFieldName(entityName, "ANALYSIS");
            string componentField = ResolveFieldName(entityName, "NAME", "COMPONENT_NAME");
            string resultTextField = ResolveFieldName(entityName, "TEXT", "RESULT_TEXT");
            string resultField = ResolveFieldName(entityName, "VALUE", "RESULT_VALUE");
            string samplingPointField = ResolveFieldName(entityName, "SAMPLING_POINT", "SAMPLE_POINT", "sampling_point");
            string loginDateFieldInResults = ResolveFieldName(entityName, "LOGIN_DATE", "LOGINDATE", "LOGIN_DT", "DATE_LOGGED");
            string unitsField = ResolveFieldName(entityName, "UNITS", "Units");

            var widthByColumnKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var captionByColumnKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var analysisByColumnKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var valuesBySample = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            var samplePointBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var sampleStatusBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var idNumericBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var loginDateBySample = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            var unitsByColumnKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var loteBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var bagBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var formBySample = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            const int dynamicColWidth = 140;

            foreach (IEntity r in resultados)
            {
                string sampleId = GetFirstNonEmpty(r, sampleIdField, "ID_TEXT", "IdText");
                string analysis = GetFirstNonEmpty(r, analysisField, "ANALYSIS", "Analysis", "ANALYSIS_NAME");
                string component = GetFirstNonEmpty(r, componentField, "COMPONENT", "Component", "COMPONENT_NAME");
                string resultText = GetFirstNonEmpty(r, resultTextField, "RESULT_TEXT", "TextResult");
                if (string.IsNullOrWhiteSpace(resultText))
                {
                    var val = r.Get(resultField);
                    resultText = val == null ? string.Empty : val.ToString();
                }

                if (string.IsNullOrEmpty(sampleId) || string.IsNullOrEmpty(component)) continue;

                var units = GetFirstNonEmpty(r, unitsField, "UNITS", "Units");
                string unitKey = string.IsNullOrWhiteSpace(units) ? string.Empty : units;
                string columnKey = component + "|" + unitKey;
                captionByColumnKey[columnKey] = component;
                if (!string.IsNullOrWhiteSpace(units)) unitsByColumnKey[columnKey] = units;
                if (!widthByColumnKey.ContainsKey(columnKey))
                {
                    widthByColumnKey[columnKey] = dynamicColWidth;
                }

                if (!valuesBySample.TryGetValue(sampleId, out var rowDict))
                {
                    rowDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    valuesBySample[sampleId] = rowDict;
                }
                rowDict[columnKey] = resultText;

                if (!string.IsNullOrEmpty(samplingPointField))
                {
                    string sp = GetFirstNonEmpty(r, samplingPointField, "SAMPLE_POINT", "SAMPLING_POINT");
                    if (!string.IsNullOrWhiteSpace(sp)) samplePointBySample[sampleId] = sp;
                }
                string st = GetFirstNonEmpty(r, "SAMPLE_STATUS", "STATUS", "SAMPLE_STATUS");
                if (!string.IsNullOrWhiteSpace(st))
                {
                    var stDesc = GetSampleStatusDescription(st);
                    sampleStatusBySample[sampleId] = string.IsNullOrWhiteSpace(stDesc) ? st : stDesc;
                }

                string idNum = GetFirstNonEmpty(r, "ID_NUMERIC", "IdNumeric", "ID_NUM");
                if (!string.IsNullOrWhiteSpace(idNum))
                {
                    idNumericBySample[sampleId] = idNum;
                }

                if (!string.IsNullOrEmpty(loginDateFieldInResults))
                {
                    var ldv = r.Get(loginDateFieldInResults);
                    if (ldv != null)
                    {
                        DateTime parsed;
                        if (DateTime.TryParse(ldv.ToString(), out parsed))
                        {
                            if (!loginDateBySample.ContainsKey(sampleId))
                                loginDateBySample[sampleId] = parsed;
                        }
                    }
                }
            }

            if (samplePointBySample.Count == 0 && idNumericBySample.Count > 0)
            {
                var idsByNum = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in idNumericBySample)
                {
                    if (!idsByNum.TryGetValue(kv.Value, out var list))
                    {
                        list = new List<string>();
                        idsByNum[kv.Value] = list;
                    }
                    list.Add(kv.Key);
                }

                string sampleEntity = "SAMPLE";
                string idNumField = ResolveFieldName(sampleEntity, "ID_NUMERIC", "IdNumeric", "ID_NUM");
                string spField = ResolveFieldName(sampleEntity, "SAMPLING_POINT", "SamplingPoint");
                if (!string.IsNullOrEmpty(idNumField) && !string.IsNullOrEmpty(spField))
                {
                    IQuery qSp = EntityManager.CreateQuery(sampleEntity);
                    var idEntities = EntityManager.CreateEntityCollection(sampleEntity);
                    foreach (var num in idsByNum.Keys)
                    {
                        var e = EntityManager.CreateEntity(sampleEntity);
                        e.Set(idNumField, num);
                        idEntities.Add(e);
                    }
                    qSp.AddIn(idNumField, idEntities, idNumField);
                    var samples = EntityManager.Select(qSp, true);
                    foreach (IEntity s in samples)
                    {
                        var num = GetFirstNonEmpty(s, idNumField);
                        var sp = GetFirstNonEmpty(s, spField);
                        if (string.IsNullOrWhiteSpace(num) || string.IsNullOrWhiteSpace(sp)) continue;
                        if (!idsByNum.TryGetValue(num, out var list)) continue;
                        foreach (var sid in list) samplePointBySample[sid] = sp;
                    }
                }
            }

            if (idNumericBySample.Count > 0)
            {
                var idsByNum2 = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in idNumericBySample)
                {
                    if (!idsByNum2.TryGetValue(kv.Value, out var list))
                    {
                        list = new List<string>();
                        idsByNum2[kv.Value] = list;
                    }
                    list.Add(kv.Key);
                }

                string sampleEntity2 = "SAMPLE";
                string idNumField2 = ResolveFieldName(sampleEntity2, "ID_NUMERIC", "IdNumeric", "ID_NUM");
                string loginField = ResolveFieldName(sampleEntity2, "DATE_LOGGED", "DATE_IN", "DATEIN", "LOGGED_ON", "DATE_LOGON");
                string loteField = ResolveFieldName(sampleEntity2, "TF_LOTE_NUMBER", "TF_LOTE_NUMERO", "TF_LOT_NUMBER", "TF_LOTE");
                string bagField = ResolveFieldName(sampleEntity2, "TF_BAG_NUMBER", "TF_BOLSA_NUMBER", "TF_BOLSA");
                string formField = ResolveFieldName(sampleEntity2, "TF_FORM_NUMBER", "TF_FORMULARIO_NUMBER", "TF_FORMULARIO");
                string spField2 = ResolveFieldName(sampleEntity2, "SAMPLING_POINT", "SamplingPoint");
                if (!string.IsNullOrEmpty(idNumField2) && !string.IsNullOrEmpty(loginField))
                {
                    IQuery qSamp = EntityManager.CreateQuery(sampleEntity2);
                    var idEntities2 = EntityManager.CreateEntityCollection(sampleEntity2);
                    foreach (var num in idsByNum2.Keys)
                    {
                        var e2 = EntityManager.CreateEntity(sampleEntity2);
                        e2.Set(idNumField2, num);
                        idEntities2.Add(e2);
                    }
                    qSamp.AddIn(idNumField2, idEntities2, idNumField2);
                    var samples2 = EntityManager.Select(qSamp, true);
                    foreach (IEntity s in samples2)
                    {
                        var num = GetFirstNonEmpty(s, idNumField2);
                        var login = !string.IsNullOrEmpty(loginField) ? s.Get(loginField) : null;
                        if (!string.IsNullOrWhiteSpace(num) && login != null)
                        {
                            DateTime parsed;
                            if (DateTime.TryParse(login.ToString(), out parsed))
                            {
                                if (idsByNum2.TryGetValue(num, out var list))
                                {
                                    foreach (var sid in list)
                                    {
                                        if (!loginDateBySample.ContainsKey(sid))
                                            loginDateBySample[sid] = parsed;
                                    }
                                }
                            }
                        }
                        if (idsByNum2.TryGetValue(num, out var listTf))
                        {
                            var lote = GetFirstNonEmpty(s, loteField);
                            var bag = GetFirstNonEmpty(s, bagField);
                            var form = GetFirstNonEmpty(s, formField);
                            foreach (var sid in listTf)
                            {
                                if (!string.IsNullOrWhiteSpace(lote)) loteBySample[sid] = lote;
                                if (!string.IsNullOrWhiteSpace(bag)) bagBySample[sid] = bag;
                                if (!string.IsNullOrWhiteSpace(form)) formBySample[sid] = form;
                            }
                        }
                        if (!string.IsNullOrEmpty(spField2))
                        {
                            var sp = GetFirstNonEmpty(s, spField2);
                            if (!string.IsNullOrWhiteSpace(sp))
                            {
                                if (idsByNum2.TryGetValue(num, out var list))
                                {
                                    foreach (var sid in list)
                                    {
                                        if (!samplePointBySample.ContainsKey(sid)) samplePointBySample[sid] = sp;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (samplePointBySample.Count > 0)
            {
                EnrichSamplePointNames(samplePointBySample);
            }

            var columnMap = new Dictionary<string, UnboundGridColumn>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in widthByColumnKey.Keys
                .OrderBy(k => captionByColumnKey.ContainsKey(k) ? captionByColumnKey[k] : string.Empty, StringComparer.OrdinalIgnoreCase))
            {
                bool tieneDatos = valuesBySample.Values.Any(row => row.TryGetValue(key, out var val) && !string.IsNullOrWhiteSpace(val));
                if (!tieneDatos) continue;
                int width = widthByColumnKey[key];
                string component = captionByColumnKey.ContainsKey(key) ? captionByColumnKey[key] : string.Empty;
                string unit = unitsByColumnKey.ContainsKey(key) ? unitsByColumnKey[key] : string.Empty;
                string caption = string.IsNullOrWhiteSpace(unit) ? component : component + " [" + unit + "]";
                var col = form.ugdPivotResults.AddColumn(key, caption, width);
                columnMap[key] = col;
            }

            foreach (var kv in valuesBySample)
            {
                var row = form.ugdPivotResults.AddRow();
                if (idNumericBySample.TryGetValue(kv.Key, out var idVal)) row.SetValue(colIdNumeric, idVal);
                row.SetValue(colSample, kv.Key);
                if (samplePointBySample.TryGetValue(kv.Key, out var spVal)) row.SetValue(colSamplePoint, spVal);
                if (loteBySample.TryGetValue(kv.Key, out var loteVal)) row.SetValue(colLote, loteVal);
                if (bagBySample.TryGetValue(kv.Key, out var bagVal)) row.SetValue(colBag, bagVal);
                if (formBySample.TryGetValue(kv.Key, out var formVal)) row.SetValue(colForm, formVal);
                if (loginDateBySample.TryGetValue(kv.Key, out var ld)) row.SetValue(colLoginDate, ld.ToString("yyyy-MM-dd HH:mm"));
                if (sampleStatusBySample.TryGetValue(kv.Key, out var stVal)) row.SetValue(colSampleStatus, stVal);
                foreach (var cell in kv.Value)
                {
                    if (columnMap.TryGetValue(cell.Key, out var col))
                    {
                        row.SetValue(col, cell.Value);
                    }
                }
            }

            if (beganUpdate) { try { grid.EndUpdate(); } catch { } }

        }


        private static int EstimateWidth(string text)
        {
            const int minWidth = 80;
            const int maxWidth = 240;
            if (string.IsNullOrWhiteSpace(text)) return minWidth;
            int approx = text.Length * 8 + 20;
            if (approx < minWidth) return minWidth;
            if (approx > maxWidth) return maxWidth;
            return approx;
        }

        private string ResolveFieldName(string entityName, params string[] candidates)
        {
            try
            {
                var table = Library.Schema.Tables[entityName];
                foreach (var c in candidates)
                {
                    try { var _ = table.Fields[c]; return c; } catch { }
                }
            }
            catch { }
            return string.Empty;
        }

        private IEntityCollection BuildResults(
            string entityName,
            string loginDateField,
            string testStatusField,
            string sampleStatusField,
            DateTime start,
            DateTime end,
            IEntityCollection samplePoints,
            bool includeCompleted)
        {
            string samplingPointField = ResolveFieldName(entityName, "SAMPLING_POINT", "SAMPLE_POINT");
            string dateCompletedField = ResolveFieldName(entityName, "DATE_COMPLETED", "COMPLETED_DATE", "DATECOMPLETED", "date_completed");
            string repControlField = ResolveFieldName(entityName, "REP_CONTROL", "rep_control");
            Action<IQuery> applyLocationFilter = (q) =>
            {
                if (samplePoints == null || samplePoints.Count == 0) return;
                if (!string.IsNullOrEmpty(samplingPointField))
                {
                    try
                    {
                        q.AddIn(samplingPointField, samplePoints, "Identity");
                        return;
                    }
                    catch
                    {

                    }
                }
                var limitedSamples = BuildSamplesFromPointsAndDates(samplePoints, start, end);
                if (limitedSamples != null && limitedSamples.Count > 0)
                {
                    q.AddIn("ID_NUMERIC", limitedSamples, "ID_NUMERIC");
                }
            };

            if (includeCompleted)
            {
                var merged = EntityManager.CreateEntityCollection(entityName);

                var qAuth = EntityManager.CreateQuery(entityName);
                if (!string.IsNullOrEmpty(testStatusField)) qAuth.AddEquals(testStatusField, "A");
                if (!string.IsNullOrEmpty(loginDateField))
                {
                    qAuth.AddGreaterThanOrEquals(loginDateField, start);
                    qAuth.AddLessThanOrEquals(loginDateField, end);
                }
                if (!string.IsNullOrEmpty(repControlField)) qAuth.AddEquals(repControlField, "REP");
                applyLocationFilter(qAuth);
                var resAuth = EntityManager.Select(qAuth, true);
                if (resAuth != null) foreach (IEntity ent in resAuth) merged.Add(ent);

                if (!string.IsNullOrEmpty(dateCompletedField))
                {
                    var qComp = EntityManager.CreateQuery(entityName);
                    if (!string.IsNullOrEmpty(sampleStatusField)) qComp.AddEquals(sampleStatusField, "C");
                    qComp.AddGreaterThanOrEquals(dateCompletedField, start);
                    qComp.AddLessThanOrEquals(dateCompletedField, end);
                    if (!string.IsNullOrEmpty(repControlField)) qComp.AddEquals(repControlField, "REP");
                    applyLocationFilter(qComp);
                    var resComp = EntityManager.Select(qComp, true);
                    if (resComp != null) foreach (IEntity ent in resComp) merged.Add(ent);
                }
                return merged;
            }
            else
            {
                var query = EntityManager.CreateQuery(entityName);
                if (!string.IsNullOrEmpty(testStatusField)) query.AddEquals(testStatusField, "A");
                if (!string.IsNullOrEmpty(loginDateField))
                {
                    query.AddGreaterThanOrEquals(loginDateField, start);
                    query.AddLessThanOrEquals(loginDateField, end);
                }
                if (!string.IsNullOrEmpty(repControlField)) query.AddEquals(repControlField, "REP");
                applyLocationFilter(query);
                return EntityManager.Select(query);
            }
        }

        private IEntityCollection BuildSamplesFromPointsAndDates(IEntityCollection samplePoints, DateTime start, DateTime end)
        {
            var samples = EntityManager.CreateEntityCollection("SAMPLE");
            if (samplePoints == null || samplePoints.Count == 0) return samples;
            string spField = ResolveFieldName("SAMPLE", "SAMPLING_POINT", "SamplingPoint");
            if (string.IsNullOrEmpty(spField)) return samples;
            string loginField = ResolveFieldName("SAMPLE", "LOGIN_DATE", "DATE_LOGGED", "DATE_IN", "DATEIN", "LOGGED_ON", "DATE_LOGON", "DATELOGGED");
            IQuery q = EntityManager.CreateQuery("SAMPLE");
            q.AddIn(spField, samplePoints, "Identity");
            if (!string.IsNullOrEmpty(loginField))
            {
                q.AddGreaterThanOrEquals(loginField, start);
                q.AddLessThanOrEquals(loginField, end);
            }
            return EntityManager.Select(q, true);
        }

        private static string GetFirstNonEmpty(IEntity entity, params string[] fields)
        {
            foreach (var f in fields)
            {
                if (string.IsNullOrEmpty(f)) continue;
                var v = entity.Get(f);
                if (v != null)
                {
                    var s = v.ToString();
                    if (!string.IsNullOrWhiteSpace(s)) return s.Trim();
                }
            }
            return string.Empty;
        }

        private string GetPhraseDescription(string phraseType, string phraseId)
        {
            var phrase = EntityManager.SelectPhrase(phraseType, phraseId);
            if (phrase != null) return phrase.ToString();
            return string.Empty;
        }



        private IEntity ResolveEntityByValue(string entityName, object valueObj)
        {
            if (valueObj == null) return null;
            var s = valueObj.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;
            return EntityManager.Select(entityName, new Identity(s));
        }

        private string GetSampleStatusDescription(string phraseId)
        {
            var types = new[] { "SAMP_STAT", "SAMP STAT", "SAMPLE_STATUS", "SAMPLE STATUS" };
            foreach (var t in types)
            {
                var d = GetPhraseDescription(t, phraseId);
                if (!string.IsNullOrWhiteSpace(d)) return d;
            }
            return string.Empty;
        }

        private void EnrichSamplePointNames(Dictionary<string, string> samplePointBySample)
        {
            string spEntity = "SAMPLE_POINT";
            string nameField = ResolveFieldName(spEntity, "POINT_NAME", "DESCRIPTION", "NAME", "DESC", "DESCRIPTION_TEXT");
            foreach (var kv in new List<KeyValuePair<string, string>>(samplePointBySample))
            {
                var spIdentity = kv.Value;
                var e = EntityManager.Select(spEntity, new Identity(spIdentity));
                if (e == null) continue;
                var label = GetFirstNonEmpty(e, nameField, "POINT_NAME", "DESCRIPTION", "NAME");
                if (!string.IsNullOrWhiteSpace(label))
                {
                    samplePointBySample[kv.Key] = label;
                }
            }
        }

        private void SafeClearBusy()
        {
            form.ClearBusy();
            form.ForceClearBusy();
        }

        private void SetHierarchyEnabled(IEntity planta, IEntity proceso, IEntity etapa)
        {
            // Aseguramos que la planta siempre esté habilitada
            if (form.pebPlanta is VisualControl vcPlanta) vcPlanta.Enabled = true;

            // Filtro para Proceso (requiere planta)
            IQuery qP = EntityManager.CreateQuery("LOCATION");
            qP.AddEquals("LocationType", "PROCESO");
            if (planta != null)
            {
                if (form.pebProceso is VisualControl vcProc) vcProc.Enabled = true;
                qP.AddEquals("ParentLocation", planta.IdentityString);
            }
            else
            {
                if (form.pebProceso is VisualControl vcProc) vcProc.Enabled = false;
                qP.AddEquals("Identity", "DUMMY_NO_DATA_XYZ"); // Evita cargar todos los procesos si no hay planta
            }
            form.pebProceso.Browse = BrowseFactory.CreateEntityBrowse(qP);

            // Filtro para Etapa (requiere proceso)
            IQuery qE = EntityManager.CreateQuery("LOCATION");
            qE.AddEquals("LocationType", "ETAPA");
            if (proceso != null)
            {
                if (form.pebEtapa is VisualControl vcEtapa) vcEtapa.Enabled = true;
                qE.AddEquals("ParentLocation", proceso.IdentityString);
            }
            else
            {
                if (form.pebEtapa is VisualControl vcEtapa) vcEtapa.Enabled = false;
                qE.AddEquals("Identity", "DUMMY_NO_DATA_XYZ"); // Evita cargar todas las etapas si no hay proceso
            }
            form.pebEtapa.Browse = BrowseFactory.CreateEntityBrowse(qE);

            // Filtro para Punto de Muestreo (requiere etapa)
            IQuery qS = EntityManager.CreateQuery("SAMPLE_POINT");
            if (etapa != null)
            {
                if (form.pebPuntoMuestreo is VisualControl vcPunto) vcPunto.Enabled = true;
                qS.AddEquals("PointLocation", etapa.IdentityString);
            }
            else
            {
                if (form.pebPuntoMuestreo is VisualControl vcPunto) vcPunto.Enabled = false;
                qS.AddEquals("Identity", "DUMMY_NO_DATA_XYZ"); // Evita cargar la pesada tabla entera de SAMPLE_POINT
            }
            form.pebPuntoMuestreo.Browse = BrowseFactory.CreateEntityBrowse(qS);
        }

        private IEntityCollection BuildSamplingPoints(IEntity plantaEntity, IEntity procesoEntity, IEntity etapaEntity, IEntity puntoEntity)
        {
            string spEntity = "SAMPLE_POINT";
            var result = EntityManager.CreateEntityCollection(spEntity);

            if (puntoEntity != null)
            {
                result.Add(puntoEntity);
                return result;
            }

            if (etapaEntity != null)
            {
                IQuery qSp = EntityManager.CreateQuery(spEntity);
                qSp.AddEquals("PointLocation", etapaEntity.IdentityString);
                return EntityManager.Select(qSp, true);
            }

            if (procesoEntity != null)
            {
                IQuery qEt = EntityManager.CreateQuery("LOCATION");
                qEt.AddEquals("LocationType", "ETAPA");
                qEt.AddAnd();
                qEt.AddEquals("ParentLocation", procesoEntity.IdentityString);
                var etapas = EntityManager.Select(qEt, true);
                if (etapas != null && etapas.Count > 0)
                {
                    IQuery qSp = EntityManager.CreateQuery(spEntity);
                    qSp.AddIn("PointLocation", etapas, "Identity");
                    return EntityManager.Select(qSp, true);
                }
                return result;
            }

            if (plantaEntity != null)
            {
                IQuery qProc = EntityManager.CreateQuery("LOCATION");
                qProc.AddEquals("LocationType", "PROCESO");
                qProc.AddAnd();
                qProc.AddEquals("ParentLocation", plantaEntity.IdentityString);
                var procesos = EntityManager.Select(qProc, true);

                if (procesos != null && procesos.Count > 0)
                {
                    IQuery qEt = EntityManager.CreateQuery("LOCATION");
                    qEt.AddEquals("LocationType", "ETAPA");
                    qEt.AddAnd();
                    qEt.AddIn("ParentLocation", procesos, "Identity");
                    var etapas = EntityManager.Select(qEt, true);
                    if (etapas != null && etapas.Count > 0)
                    {
                        IQuery qSp = EntityManager.CreateQuery(spEntity);
                        qSp.AddIn("PointLocation", etapas, "Identity");
                        return EntityManager.Select(qSp, true);
                    }
                }
                return result;
            }

            return null;
        }


    }
}
