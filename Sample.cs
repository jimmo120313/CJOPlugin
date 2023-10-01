using Microsoft.EngagementInsights;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.Remoting.Services;
using System.ServiceModel;

namespace Sce.Solution.TriggerConditionalApprovalCJO
{

    internal class CJOTriggers
    {
        public string analyticsGuid { get; set; }
        public string cjoName { get; set; }
    }
    
    public class TriggerConditionalApprovalCJOPlugin : IPlugin
    {
        public async void Execute(IServiceProvider serviceProvider)
        {

            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {

                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    //tracingService.Trace("Started");
                    var newOpportunity = (Entity)context.InputParameters["Target"];
                    Money newAmount = newOpportunity.Contains("sce_approved_amount")?(Money)newOpportunity["sce_approved_amount"]:null;
                    OptionSetValue newStatus = newOpportunity.Contains("sce_applicationstatus") ?(OptionSetValue)newOpportunity["sce_applicationstatus"]:null;
                    //tracingService.Trace("Get Opportunity");
                    Entity existOpportunity = (Entity)context.PreEntityImages["opportunity"];
                    NotificationType notificationType = GetNotificationType(newAmount, newStatus,existOpportunity);

                    if (notificationType!=NotificationType.None)
                    {
                        //tracingService.Trace("before get All lead");
                        EntityCollection allLeads = GetApplicationLeads(context.PrimaryEntityId, service);
                        //tracingService.Trace("before trigger CJO");
                        TriggerCJO(notificationType, allLeads,existOpportunity, service, tracingService);
                    }
                    
       
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in ConditionalApprovalTrigerCJO Plugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("ConditionalApprovalTrigerCJO: {0}", ex.ToString());
                    throw;
                }
            }
        }

        private EntityCollection GetApplicationLeads(Guid oppoId, IOrganizationService svc)
        {
            var query = new QueryExpression("lead")
            {
                ColumnSet = new ColumnSet("fullname"),
                Criteria = new FilterExpression(LogicalOperator.And),
                TopCount = 10
            };
            query.Criteria.AddCondition("qualifyingopportunityid", ConditionOperator.Equal, oppoId);
            query.AddOrder("fullname", OrderType.Ascending);

            EntityCollection results = svc.RetrieveMultiple(query);
            return results;
        }

        private void TriggerCJO(NotificationType notificationType, EntityCollection leads,Entity oppo, IOrganizationService svc, ITracingService trace)
        {
            trace.Trace("TriggerCJO started");
            CJOTriggers cjoDetail = GetCJOTriggersDetail(svc, "sce__ConditionalApproveCJODetail");
            Analytics analytics = new Analytics(cjoDetail.analyticsGuid);

            foreach (var lead in leads.Entities)
            {
                User leadUser = new User()
                {
                    UserAuthId = lead.Id.ToString()
                };

                analytics.SetUser(leadUser);
                Event approveEvent = new Event(cjoDetail.cjoName);
                //trace.Trace("setProperty");

                approveEvent.SetProperty("datevariable", DateTime.Now);
                approveEvent.SetProperty("opportunity", oppo.Id);
                approveEvent.SetProperty("textvariable2", "N/A");
                approveEvent.SetProperty("bindingid", lead.Id);
                approveEvent.SetProperty("textvariable3", "N/A");
                approveEvent.SetProperty("textvariable4", "N/A");
                approveEvent.SetProperty("textvariable1", "N/A");

                analytics.TrackEvent(approveEvent);

                analytics.UploadNow();
            }
        }

        private CJOTriggers GetCJOTriggersDetail(IOrganizationService svc,string variableName)
        {
            Dictionary<string, string> envVariables = new Dictionary<string, string>();

            var query = new QueryExpression("environmentvariabledefinition")
            {
                Criteria = new FilterExpression(LogicalOperator.And),
                ColumnSet = new ColumnSet("statecode", "defaultvalue", "valueschema",
                  "schemaname", "environmentvariabledefinitionid", "type"),
                LinkEntities =
                        {
                            new LinkEntity
                            {
                                JoinOperator = JoinOperator.LeftOuter,
                                LinkFromEntityName = "environmentvariabledefinition",
                                LinkFromAttributeName = "environmentvariabledefinitionid",
                                LinkToEntityName = "environmentvariablevalue",
                                LinkToAttributeName = "environmentvariabledefinitionid",
                                Columns = new ColumnSet("statecode", "value", "environmentvariablevalueid"),
                                EntityAlias = "v"
                            }
                        }
            };

            query.Criteria.AddCondition("schemaname", ConditionOperator.Equal, variableName); 
            var results = svc.RetrieveMultiple(query);
            if (results?.Entities.Count > 0)
            {
                foreach (var entity in results.Entities)
                {
                    var schemaName = entity.GetAttributeValue<string>("schemaname");
                    var value = entity.GetAttributeValue<AliasedValue>("v.value")?.Value?.ToString();
                    var defaultValue = entity.GetAttributeValue<string>("defaultvalue");

                    if (schemaName != null && !envVariables.ContainsKey(schemaName))
                        envVariables.Add(schemaName, string.IsNullOrEmpty(value) ? defaultValue : value);
                }
            }

            CJOTriggers triggerDetail = JsonConvert.DeserializeObject<CJOTriggers>(envVariables[variableName]);
            
            return triggerDetail;
        
        }

        private NotificationType GetNotificationType(Money newAmount, OptionSetValue status, Entity existOpportunity)
        {

            Money oldAmount = existOpportunity.Contains("sce_financial_conditionally_approved_amount") ? (Money)existOpportunity["sce_financial_conditionally_approved_amount"] : null;
            var cstatus = status==null ? (OptionSetValue)existOpportunity["sce_applicationstatus"] : status;
            //Conditionally Approved = 8
            //Pending Conditionally Approved = 9
            //Conditionally Declined = 10
            //Approved = 3
            //Accepted = 4
            if ( cstatus.Value == 8)
            {
                if (newAmount.Value.Equals(oldAmount.Value) || newAmount==null)
                {
                    return NotificationType.None;
                }
                else
                {
                   return  oldAmount ==null?NotificationType.ConditionalApproved : NotificationType.ConditionalApprovedUpdated;
                }
            }
            else
            {
                return NotificationType.None;
            }
        }
    }
}