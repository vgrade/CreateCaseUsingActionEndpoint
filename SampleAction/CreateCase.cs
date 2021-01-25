using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace CreateCaseEndpoint
{
    public class CreateCase : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext pluginContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationService service = serviceFactory.CreateOrganizationService(null);

            if (pluginContext.MessageName.ToLower() == "new_createcaseaction")
            {
                try
                {
                    string inputJson = (string)pluginContext.InputParameters["Input"];
                    CaseModel inputData = JsonConvert.DeserializeObject<CaseModel>(inputJson);
                    Guid contactGuid = GetContactRecord(service,inputData);
                    Guid CaseGuid = CreateCaseRecord(service, inputData, contactGuid);
                    pluginContext.OutputParameters["Result"] = CaseGuid.ToString();
                    tracing.Trace("Plugin executed successfully");
                }
                catch (InvalidPluginExecutionException ex)
                {
                    throw new InvalidPluginExecutionException("An error occured" + ex.Message);
                }
            }
        }

        public Guid GetContactRecord(IOrganizationService service, CaseModel inputData)
        {
            Guid contactGuid;
            QueryExpression qe = new QueryExpression("contact");
            qe.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, inputData.MobileNumber);

            Entity entity = new Entity("contact");
            entity["firstname"] = inputData.FirstName;
            entity["lastname"] = inputData.LastName;
            entity["emailaddress1"] = inputData.EmailAddress;
            entity["birthdate"] = Convert.ToDateTime(inputData.Dob);
            entity["fax"] = inputData.Fax;
            entity["gendercode"] = inputData.Gender.ToUpper() == "M" ? new OptionSetValue(1) : new OptionSetValue(2);

            EntityCollection enColl = service.RetrieveMultiple(qe);
            if (enColl.Entities.Count > 0)
            {
                contactGuid = enColl.Entities[0].Id;
                entity.Id = enColl.Entities[0].Id;
                service.Update(entity);
            }
            else
            {
                entity["mobilephone"] = inputData.MobileNumber;
                contactGuid = service.Create(entity);
            }
            return contactGuid;
        }

        public Guid CreateCaseRecord(IOrganizationService service, CaseModel inputData, Guid ContactGuid)
        {
            Entity entity = new Entity("incident");
            entity["title"] = inputData.CaseTitle;
            entity["customerid"] = new EntityReference("contact", ContactGuid);
            entity["prioritycode"] = inputData.Priority.ToUpper() == "HIGH" ? new OptionSetValue(1) : inputData.Gender.ToUpper() == "NORMAL" ? new OptionSetValue(2) : new OptionSetValue(3);
            return service.Create(entity);
        }
    }
}
