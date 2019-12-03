using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Extending.Dyn365.v9.CustomWorkflowActivity
{
    /// <summary>
    /// Post message to Azure Service Bus 
    /// </summary>
    public sealed class PostMessageToASB : CodeActivity
    {
        #region members

        ITracingService tracingService = null;

        #endregion

        #region Custom Workflow Activity Parameters

        [RequiredArgument]
        [Input("Service Endpoint Id")]
        [ReferenceTarget("serviceendpoint")]
        public InArgument<EntityReference> serviceEndPoint { get; set; }

        [Input("Properties Set")]
        public InArgument<string> propertiesSet { get; set; }

        [Input("Custom Message")]
        public InArgument<string> customMessage { get; set; }

        [Default("false")]
        [Output("Is Successful Execution")]
        public OutArgument<bool> isSuccessfulExecution { get; set; }

        [RequiredArgument]
        [Output("Error Message")]
        public OutArgument<string> errorMessage { get; set; }

        #endregion 

        protected override void Execute(CodeActivityContext context)
        {
            Entity targetEntity = null;

            //define required services 
            IWorkflowContext workflowcontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowcontext.InitiatingUserId);
            IServiceEndpointNotificationService cloudService = context.GetExtension<IServiceEndpointNotificationService>();
            tracingService = context.GetExtension<ITracingService>();

            //extract input parameters 
            EntityReference serviceEndPoint = this.serviceEndPoint.Get(context);
            string propertiesSet = this.propertiesSet.Get(context);
            string customMessage = this.customMessage.Get(context);

            tracingService.Trace("Execution start");

            try
            {
                //clear redundant data from context to reduce payload weight
                workflowcontext.PreEntityImages.Clear();
                workflowcontext.PostEntityImages.Clear();
                workflowcontext.InputParameters.Clear();
                workflowcontext.OutputParameters.Clear();
                workflowcontext.SharedVariables.Clear();

                //extract required entity attributes if specified 
                if (!string.IsNullOrEmpty(propertiesSet))
                {
                    string[] attributes = propertiesSet.Split(';');

                    ColumnSet cs = attributes.Length > 0 ? new ColumnSet(attributes) : new ColumnSet(propertiesSet);

                    //retrieve target record attributes 
                    targetEntity = service.Retrieve(workflowcontext.PrimaryEntityName, workflowcontext.PrimaryEntityId, cs);
                }

                //parse retrieved attributes into SharedVariables collection
                if (targetEntity != null)
                {
                    foreach (var attribute in targetEntity.Attributes)
                    {
                        workflowcontext.SharedVariables.Add(attribute.Key, attribute.Value);
                    }
                }

                //add custom message to the SharedVariables collection
                if (!string.IsNullOrEmpty(customMessage))
                {
                    workflowcontext.SharedVariables.Add("customMessage", customMessage);
                }

                //send message to target service endpoint 
                tracingService.Trace("Starting posting the execution context");

                string response = cloudService.Execute(new EntityReference("serviceendpoint",
                    serviceEndPoint.Id),
                    workflowcontext);

                if (!String.IsNullOrEmpty(response))
                {
                    tracingService.Trace("Response = {0}", response);
                }

                tracingService.Trace("completed posting the execution context");
            }
            catch (Exception ex)
            {
                //handle exception
                tracingService.Trace("Exception: {0}", ex.Message);
                //set output parameter 
                errorMessage.Set(context, ex.Message);
            }

            //set output parameter 
            isSuccessfulExecution.Set(context, true);

            tracingService.Trace("Execution end");
        }
    }
}

