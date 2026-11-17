using DevExpress.XtraSpreadsheet.Import.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aucotec.EngineeringBase.Interop;
using EBApp = Aucotec.EngineeringBase.Interop.Application;
using DevExpress.XtraSpreadsheet.UI;

namespace Reefer_Container
{
    public class EBHelper
    {
        private int sleep = 1;
        public EBHelper()
        {
            Form1.ebapp = GetEB();
        }

        public EBApp GetEB()
        {
            try
            {
                string progID = "EngineeringBase.Application.31"; // 버전에 맞게 조정 필요
                Type type = Type.GetTypeFromProgID(progID);
                if (type == null)
                {
                    Debug.WriteLine($"❌ EB ProgID '{progID}' not found. Check version or registration.");
                    return null;
                }

                object dynamicEB = Activator.CreateInstance(type);
                EBApp ebapp = dynamicEB as EBApp;

                Debug.WriteLine($"{ebapp.Name}");

                return ebapp;
            }
            catch
            {

            }

            return null;

        }

        // aucAttrDesignation -> 텍스트로만 찾기
        public Filter Get_filter(string search_text)
        {
            Filter filter = new Filter();
            filter.Add(AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, search_text);

            return filter;
        }
        // ObjecType추가
        public Filter Get_filter(AucType type, string search_text = "")
        {
            Filter filter = new Filter();
            filter.TypeID = type;
            filter.Add(AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, search_text);

            return filter;
        }
        // ObjectKind추가
        public Filter Get_filter(AucObjectKind objectKind, string search_text = "")
        {
            Filter filter = new Filter();
            filter.Kind = objectKind;
            filter.Add(AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, search_text);

            return filter;
        }
        public Filter Get_filter(AucObjectKind objectKind, AucType type, string search_text = "")
        {
            Filter filter = new Filter();
            filter.Kind = objectKind;
            filter.TypeID = type;
            filter.Add(AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, search_text);

            return filter;
        }

        public string ReadAttributeValue(IAucObject obj, AucAttribute attr)
        {
            try
            {
                return obj.Attributes.ItemByID(attr).Value.ToString();
            }
            catch { return ""; }
        }
        
        public string GetFormulaResult(IAucObject obj, string formula)
        {
            //A + att id + ;
            string value = "";
            try
            {
                obj.ExecuteFormula(formula, out value);
            }
            catch
            {
                return "";
            }
            return value;
        }

        public string GetFormula(IAucObject obj, object value)
        {
            try
            {
                string sformula = $"El2;A{value};";

                obj.ExecuteFormula(sformula, out string sResult);
                //Debug.WriteLine($"GetFormula : {sResult}");
                return sResult;
            }
            catch
            {
                return "";
            }

        }
        public IAucObject Get_object(IAucObject oParent, string value,
                                            AucObjectKind kind = AucObjectKind.aucObjDevice,
                                            AucType type = AucType.aucTypeUnspecified,
                                            AucVbSearchType search = AucVbSearchType.aucSearchFlat,
                                            AucAttribute aid = AucAttribute.aucAttrDesignation,
                                            AucVbFindCondition find = AucVbFindCondition.aucCondEqual,
                                            bool bCreate = true)
        {
            IAucObject oRt = null;
            IAucObjects oFinds = oParent.FindObjects(kind, search, type, aid, find, value);

            if (oFinds.Count > 0)
            {
                oRt = oFinds.Item(1);
            }
            else
            {
                if (bCreate)
                {
                    oRt = oParent.Children.AddNew(kind, type);
                    oRt.Attributes.ItemByID(aid).Value = value;
                    oRt.Store();
                }
                else { oRt = null; }
            }

            return oRt;
        }

        public bool SetAttribute(IAucObject obj, int attr_id, string value)
        {
            try
            {
                bool has = TryFindAttribute(obj, attr_id);
                if (!has)
                {
                    has = AddAttribute(obj, attr_id, "Operating Data");
                    if (has)
                    {
                        obj.Attributes.ItemByID((AucAttribute)attr_id).Value = value;
                        obj.Store();
                        return true;
                    }
                    else
                        return false;
                }
                else
                {
                    obj.Attributes.ItemByID((AucAttribute)attr_id).Value = value;
                    obj.Store();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        public bool SetAttribute(IAucObject obj, int attr_id, int value)
        {
            try
            {
                bool has = TryFindAttribute(obj, attr_id);
                if (!has)
                {
                    has = AddAttribute(obj, attr_id, "Operating Data");
                    if (has)
                    {
                        obj.Attributes.ItemByID((AucAttribute)attr_id).Value = value.ToString();
                        obj.Store();
                        return true;
                    }
                    else
                        return false;
                }
                else
                {
                    obj.Attributes.ItemByID((AucAttribute)attr_id).Value = value.ToString();
                    obj.Store();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool SetAttributes(IAucObject obj, IEnumerable<KeyValuePair<int, string>> attributes)
        {
            if (obj == null || attributes == null)
                return false;

            try
            {
                var attributeValues = attributes
                    .Where(x => x.Key > 0)
                    .GroupBy(x => x.Key)
                    .Select(x => x.Last())
                    .ToList();

                if (attributeValues.Count == 0)
                    return false;

                bool hasChanges = false;
                foreach (var attributeValue in attributeValues)
                {
                    int attrId = attributeValue.Key;
                    string value = attributeValue.Value ?? string.Empty;

                    if (!TryFindAttribute(obj, attrId))
                    {
                        obj.Attributes.Add((AucAttribute)attrId, string.Empty, "Operating Data");
                        hasChanges = true;
                    }

                    IAucAttribute attribute = obj.Attributes.ItemByID((AucAttribute)attrId);
                    string currentValue = attribute.Value?.ToString() ?? string.Empty;
                    if (string.Equals(currentValue, value, StringComparison.Ordinal))
                        continue;

                    attribute.Value = value;
                    hasChanges = true;
                }

                if (hasChanges)
                    obj.Store();

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SetAttributes failed: {ex.Message}");
                return false;
            }
        }

        public bool AddAttribute(IAucObject obj, int attr_id, string tab_name, string value = "")
        {
            try
            {
                bool has = TryFindAttribute(obj, attr_id);
                if (!has)
                {
                    obj.Attributes.Add((AucAttribute)attr_id, value, tab_name);
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
        public bool AddSourceAsociation(IAucObject target, IAucObject source, int role)
        {
            try
            {
                target.SourceAssociations.New((AucRole)role, source);
                Debug.WriteLine($"OK_AddSourceAsociation: ({role})  s:{source.Name} -> t:{target.Name}");
                System.Threading.Thread.Sleep(sleep);
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Err Source: ({e.Message})");
                return false;
            }

        }
        public bool AddTargetAsociation(IAucObject target, IAucObject source, int role)
        {
            try
            {
                target.TargetAssociations.New((AucRole)role, source);
                Debug.WriteLine($"OK_AddTargetAsociation: ({role})  t:{target.Name} <- s:{source.Name}");
                System.Threading.Thread.Sleep(sleep);
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Err Target: ({e.Message})");
                return false;
            }

        }
        public void AssociationClear(IAucObject item, int role)
        {
            if (item.SourceAssociations.Count > 0)
            {
                for (int i = item.SourceAssociations.Count; i >= 1; i--)
                {
                    try
                    {
                        item.SourceAssociations.Delete((AucRole)role, (IAucObject)item.SourceAssociations.Item(i).RelatedObject);
                        System.Threading.Thread.Sleep(sleep);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"{ex.Message}");

                    }

                }
            }

            if (item.TargetAssociations.Count > 0)
            {
                for (int i = item.TargetAssociations.Count; i >= 1; i--)
                {
                    try
                    {
                        item.TargetAssociations.Item(i).RelatedObject.SourceAssociations.Delete((AucRole)role, (IAucObject)item);
                        System.Threading.Thread.Sleep(sleep);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"{ex.Message}");

                    }

                }
            }

        }
        public bool AutoNewAssociation_reefer(IAucObject reefer, IAucObject newGangbox, int role = 119)
        {
            if (reefer == null || newGangbox == null) return false;

            try
            {
                // 새로운 Gangbox와 reefer 연결
                newGangbox.SourceAssociations.New((AucRole)role, reefer);
                System.Threading.Thread.Sleep(sleep);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"{ex.Message}");
                return false;
            }           

            return true;
        }
        public bool AutoDelAssociation_reefer(IAucObject reefer, IAucObject oldGangbox, int role = 119)
        {
            if (reefer == null || oldGangbox == null) return false;

            try
            {
                // 현재 Gangbox에서 reefer는 끊고
                oldGangbox.SourceAssociations.Delete((AucRole)role, reefer);
                System.Threading.Thread.Sleep(sleep);

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"{ex.Message}");
                return false;
            }

            return true;
        }
        /// <summary>
        /// 실제 EB Association을 기준으로 GangBox(Source) -> Reefer(Target) 관계를 동기화한다.
        /// </summary>
        public bool SynchronizeGangBoxReeferAssociations(IAucObject gangbox,IEnumerable<IAucObject> desiredReefers,int role,
            out int addedCount,out int removedCount)
        {
            List<IAucObject> detachedReefers;
            return SynchronizeGangBoxReeferAssociations(
                gangbox,
                desiredReefers,
                role,
                out addedCount,
                out removedCount,
                out detachedReefers);
        }

        public bool SynchronizeGangBoxReeferAssociations(IAucObject gangbox,IEnumerable<IAucObject> desiredReefers, int role
            ,out int addedCount,out int removedCount,out List<IAucObject> detachedReefers)
        {
            addedCount = 0;
            removedCount = 0;
            detachedReefers = new List<IAucObject>();

            if (gangbox == null || desiredReefers == null)
                return false;

            try
            {
                var desiredById = desiredReefers
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.ID))
                    .GroupBy(x => x.ID)
                    .ToDictionary(x => x.Key, x => x.First());

                var currentReefers = new List<IAucObject>();
                foreach (IAucAssociation association in gangbox.SourceAssociations.Filter((AucRole)role))
                {
                    IAucObject related = association.RelatedObject;
                    if (related != null)
                        currentReefers.Add(related);
                }

                // 기존 GangBox에는 있지만 최종 Excel 목록에서 빠진 Reefer만 반환한다.
                // 계속 사용하는 Reefer는 아래에서 재연결되므로 Position을 초기화하면 안 된다.
                detachedReefers = currentReefers
                    .Where(x => x != null &&
                                !string.IsNullOrWhiteSpace(x.ID) &&
                                !desiredById.ContainsKey(x.ID))
                    .GroupBy(x => x.ID)
                    .Select(x => x.First())
                    .ToList();

                // 최종 Excel에서 빠진 관계만 삭제한다.
                foreach (IAucObject currentReefer in detachedReefers)
                {
                    gangbox.SourceAssociations.Delete((AucRole)role, currentReefer);
                    removedCount++;
                    System.Threading.Thread.Sleep(sleep);
                }

                // Snapshot이 아니라 실제 EB Association을 기준으로 판단한다.
                // 목표 GangBox에 이미 연결되어 있더라도 다른 GangBox와의 중복 연결은 정리한다.
                foreach (IAucObject reefer in desiredById.Values)
                {
                    var oldGangBoxes = new List<IAucObject>();
                    bool alreadyAssociated = false;

                    foreach (IAucAssociation association in reefer.TargetAssociations.Filter((AucRole)role))
                    {
                        IAucObject related = association.RelatedObject;
                        if (related == null)
                            continue;

                        if (string.Equals(related.ID, gangbox.ID, StringComparison.OrdinalIgnoreCase))
                            alreadyAssociated = true;
                        else
                            oldGangBoxes.Add(related);
                    }

                    foreach (IAucObject oldGangBox in oldGangBoxes
                        .Where(x => x != null && !string.IsNullOrWhiteSpace(x.ID))
                        .GroupBy(x => x.ID, StringComparer.OrdinalIgnoreCase)
                        .Select(x => x.First()))
                    {
                        oldGangBox.SourceAssociations.Delete((AucRole)role, reefer);
                        removedCount++;
                        System.Threading.Thread.Sleep(sleep);
                    }

                    if (!alreadyAssociated)
                    {
                        gangbox.SourceAssociations.New((AucRole)role, reefer);
                        addedCount++;
                        System.Threading.Thread.Sleep(sleep);
                    }
                }

                Debug.WriteLine(
                    $"SynchronizeGangBoxReeferAssociations: {gangbox.Name}, " +
                    $"added={addedCount}, removed={removedCount}, desired={desiredById.Count}, " +
                    $"detached={detachedReefers.Count}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SynchronizeGangBoxReeferAssociations failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 실제 EB Association을 기준으로 RDB(Source) -> GangBox(Target) 관계를 동기화한다.
        /// </summary>
        public bool SynchronizeRdbGangBoxAssociation(IAucObject rdb,IAucObject gangbox,int role = 119)
        {           
            if (rdb == null || gangbox == null) return false;

            try
            {
                bool alreadyAssociated = false;
                var oldRdbs = new List<IAucObject>();

                foreach (IAucAssociation associ in gangbox.TargetAssociations.Filter((AucRole)role))
                {
                    IAucObject related = associ.RelatedObject;
                    if (related == null)
                        continue;

                    if (related.ID == rdb.ID)
                        alreadyAssociated = true;
                    else
                        oldRdbs.Add(related);
                }

                foreach (IAucObject oldRdb in oldRdbs)
                {
                    oldRdb.SourceAssociations.Delete((AucRole)role, gangbox);
                    System.Threading.Thread.Sleep(sleep);
                }

                if (!alreadyAssociated)
                {
                    rdb.SourceAssociations.New((AucRole)role, gangbox);
                    System.Threading.Thread.Sleep(sleep);
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"{ex.Message}");
                return false;
            }
        }
        
        public string GetFullID(string shortID)
        {
            if (shortID.Length < 8) return "";
            return "00000000-0000-0000-0000-0000" + shortID.Substring(shortID.Length - 8);
        }

        public IAucObject GetObjectByID(string id)
        {
            return Form1.ebapp.Utils.GetSnglObjectByID(id);
        }

        public bool TryFindAttribute(IAucObject obj, int attr_id)
        {
            try
            {                
                IAucAttribute attr_obj = obj.Attributes.ItemByID((AucAttribute)attr_id);
                if (attr_obj == null)
                    return false;
                else
                    return true;
            }
            catch
            {
                return false;
            }
        }
        public bool TryMoveTo(IAucObject Target, IAucObject obj)
        {
            if (Target == null)
                return false;
            if (obj == null)
                return false;

            try
            {
                bool tnf = obj.MoveTo(Target);  
                return tnf;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TryMoveTo err: {obj.Name}->{Target.Name}\r{ex.Message}");                
                return false;
            }

        }
        public IAucObject TryCopyTo(IAucObject Parent, IAucObject obj, bool children = true)
        {
            if (Parent == null)
                return null;
            if (obj == null)
                return null;

            try
            {
                IAucObject tnf = obj.CopyTo(Parent, children);
                return tnf;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TryCopyTo err: {obj.Name}->{Parent.Name}\r{ex.Message}");
                return null;
            }

        }
        public IAucObject TryAddNew(IAucObject Parent, AucObjectKind kind, AucType type)
        {
            if (Parent == null)
                return null;
           
            try
            {
                IAucObject tnf = Parent.Children.AddNew(kind, type);
                return tnf;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TryAddNew err: {Parent.Name}\r{ex.Message}");                
                return null;
            }

        }



    }
}
