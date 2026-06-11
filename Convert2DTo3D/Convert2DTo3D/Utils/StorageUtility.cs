using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Convert2DTo3D.Utils
{
    public class StorageUtility
    {
        public static Schema CreateSchema(Guid guid, string fieldName, Type type)
        {
            try
            {
                if (SchemaBuilder.GUIDIsValid(guid))
                {
                }

                SchemaBuilder schemaBuilder = new SchemaBuilder(guid);

                if (schemaBuilder.AcceptableName(fieldName))
                {
                }

                if (type == typeof(int))
                    schemaBuilder.SetSchemaName("Mark_IntSchema");
                else // if (type == typeof(string))
                    schemaBuilder.SetSchemaName("Mark_StrSchema");

                // Have to define the field name as string and
                // set the type using typeof method

                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(fieldName, type);
                return schemaBuilder.Finish();
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static object GetValue(Element element, Schema schema, string storageName, Type type)
        {
            try
            {
                if (element == null)
                    return null;
                var entity = element.GetEntity(schema);
                if (entity == null)
                    return null;

                object value = null;
                if (type == typeof(int))
                    value = entity.Get<int>(storageName);
                else
                    value = entity.Get<string>(storageName);

                if (value == null)
                {
                    return null;
                }
                return value;
            }
            catch (System.Exception ex)
            {
            }
            return null;
        }

        public static bool SetValue(Element element, Guid guid, string fieldName, Type type, object value)
        {
            if (element == null)
                return false;

            try
            {
                var schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(guid);
                if (schema == null)
                    return false;

                var entity = element.GetEntity(schema);
                if (entity == null || entity.Schema == null)
                    return false;

                if (type == typeof(IList<ElementId>))
                {
                    var list = (IList<ElementId>)value;
                    entity.Set<IList<ElementId>>(fieldName, list);
                }
                else if (type == typeof(string))
                {
                    var str = (string)value;
                    entity.Set<string>(fieldName, str);
                }
                else if (type == typeof(ElementId))
                {
                    var id = (ElementId)value;
                    entity.Set<ElementId>(fieldName, id);
                }
                else if (type == typeof(int))
                {
                    var iValue = (int)value;
                    entity.Set<int>(fieldName, iValue);
                }
                else if (type == typeof(bool))
                {
                    var iValue = (bool)value == true ? 1 : 0;
                    entity.Set<int>(fieldName, iValue);
                }
                else if (type == typeof(double))
                {
                    var iValue = (double)value;
#if DEBUG2019 || DEBUG_2020 || DEBUG2021 || RELEASE2019 || RELEASE2020 || RELEASE2021 || RELEASE2019_VN || RELEASE2020_VN || RELEASE2021_VN
                    entity.Set<double>(fieldName, iValue, DisplayUnitType.DUT_CUSTOM);
#else
                    entity.Set<double>(fieldName, iValue, UnitTypeId.Custom);
#endif
                }

                element.Document.Regenerate();

                element.SetEntity(entity);

                return true;
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;
                return false;
            }
        }

        public static bool AddEntity(Element element, Guid guid, string name, object value)
        {
            if (element == null)
                return false;

            try
            {
                Type type = value.GetType();
                var schema = Schema.Lookup(guid);
                if (schema == null)
                    schema = StorageUtility.CreateSchema(guid, name, type);

                var entity = new Autodesk.Revit.DB.ExtensibleStorage.Entity(schema);

                if (type == typeof(int))
                    entity.Set(name, (int)value);
                else
                    entity.Set(name, (string)value);

                element.SetEntity(entity);

                return true;
            }
            catch (System.Exception ex)
            {
            }

            return false;
        }
    }

    public static class UtDataStorage
    {
        public static Schema CreateSchema<T>(string dataStorageName, List<string> listParametersName)
        {
            var guide = GenerateGuidFromText(dataStorageName);
            var schema = Schema.Lookup(guide);
            if (schema != null) return schema;

            var schemaBuilder = new SchemaBuilder(guide);
            schemaBuilder.SetSchemaName(dataStorageName);
            foreach (var key in listParametersName)
            {
                var fieldTypeHasUnit = typeof(T) == typeof(double) || typeof(T) == typeof(XYZ);
                if (fieldTypeHasUnit)
                    schemaBuilder.AddSimpleField(key, typeof(T)).SetSpec(SpecTypeId.Custom);
                else schemaBuilder.AddSimpleField(key, typeof(T));
            }
            return schemaBuilder.Finish();
        }

        public static T GetElementParameterDataStorage<T>(this Element element, string dataStorageName, string keyParameter)
        {
            var guide = GenerateGuidFromText(dataStorageName);
            var schema = Schema.Lookup(guide);
            if (schema == null) return default;
            var createdInfoEntity = element.GetEntity(schema);
            if (createdInfoEntity.Schema == null) return default;
            if (typeof(T) == typeof(double) || typeof(T) == typeof(double?)
                || typeof(T) == typeof(int) || typeof(T) == typeof(int?)
                || typeof(T) == typeof(bool) || typeof(T) == typeof(bool?)
                || typeof(T) == typeof(string) || typeof(T) == typeof(ElementId)
                || typeof(T) == typeof(XYZ) || typeof(T) == typeof(Guid) || typeof(T) == typeof(Entity))
            {
                var fieldTypeHasUnit = typeof(T) == typeof(double) || typeof(T) == typeof(XYZ);
                return fieldTypeHasUnit ? createdInfoEntity.Get<T>(keyParameter, UnitTypeId.Custom) : createdInfoEntity.Get<T>(keyParameter);
            }
            var data = createdInfoEntity.Get<string>(keyParameter);
            return JsonConvert.DeserializeObject<T>(data);
        }

        public static void SetElementParameterDataStorage<T>(this Element element, string dataStorageName, string keyParameter, T parameterValue)
        {
            if (parameterValue == null)
            {
                Debug.Assert(false, "Parameter value is null.Cant not set to storage");
                return;
            }
            var guide = GenerateGuidFromText(dataStorageName);
            var schema = Schema.Lookup(guide);
            object data = parameterValue;
            bool inValidType = false;
            if (typeof(T) != typeof(double) && typeof(T) != typeof(double?) && typeof(T) != typeof(int) && typeof(T) != typeof(int?) && typeof(T) != typeof(bool) &&
                 typeof(T) != typeof(bool?) && typeof(T) != typeof(string) && typeof(T) != typeof(ElementId) && typeof(T) != typeof(XYZ) &&
                 typeof(T) != typeof(Guid) && typeof(T) != typeof(Entity))
            {
                data = JsonConvert.SerializeObject(parameterValue);
                inValidType = true;
            }
            if (schema != null)
            {
                Entity createdInfoEntity = element.GetEntity(schema);
                if (createdInfoEntity.Schema != null)
                {
                    if (inValidType) createdInfoEntity.Set<string>(keyParameter, $"{data}");
                    else createdInfoEntity.Set<T>(keyParameter, (T)data);
                    element.SetEntity(createdInfoEntity);
                    return;
                }
            }

            Entity storageEntity;

            if (inValidType)
            {
                storageEntity = new Entity(CreateSchema<string>(dataStorageName, new List<string> { keyParameter }));
                storageEntity.Set(keyParameter, $"{data}");
            }
            else
            {
                storageEntity = new Entity(CreateSchema<T>(dataStorageName, new List<string> { keyParameter }));

                var fieldTypeHasUnit = typeof(T) == typeof(double) || typeof(T) == typeof(XYZ);
                if (fieldTypeHasUnit)
                    storageEntity.Set<T>(keyParameter, (T)data, UnitTypeId.Custom);
                else
                    storageEntity.Set<T>(keyParameter, (T)data);
            }

            element.SetEntity(storageEntity);
        }

        public static Guid GenerateGuidFromText(string input)
        {
            using var sha256 = SHA256.Create();
            byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
            return new Guid(hash.Take(16).ToArray());
        }
    }
}