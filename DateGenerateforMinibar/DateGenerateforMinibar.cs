using Beezlabs.RPAHive.Lib;
using Beezlabs.RPAHive.Lib.Models;
using System;

namespace Beezlabs.RPA.Bots
{
    public class DateGenerateforMinibar : RPABotTemplate
    {
        string StartDate;
        string Endate;
        DateTime Startdate =new DateTime(); 
        DateTime EndTimee=new DateTime();
        VariableModel VarInput = null;
        protected override void BotLogic(BotExecutionModel botExecutionModel)
        {
            //System.Diagnostics.Debugger.Launch();
          //  System.Diagnostics.Debugger.Break();

            try {


                VarInput = null;
                botExecutionModel.proposedBotInputs.TryGetValue("startDate", out VarInput);
                if (VarInput != null && VarInput.value != null && VarInput.value != "")
                {
                    DateTime d = Convert.ToDateTime((string)VarInput.value);
                    StartDate = d.ToString("yyyy-MM-dd");
                    LogMessage(this.GetType().Name, "start date is present");
                }
                else
                {
                    LogMessage(this.GetType().FullName, "start Date is missing");
                }
                VarInput = null;
                botExecutionModel.proposedBotInputs.TryGetValue("endDate", out VarInput);
                if (VarInput != null && VarInput.value != null && VarInput.value != "")
                {
                    DateTime d1 = Convert.ToDateTime((string)VarInput.value);
                    Endate = d1.ToString("yyyy-MM-dd");
                    LogMessage(this.GetType().Name, "end date is present");
                }
                else
                {
                    LogMessage(this.GetType().Name, "end date is missing");
                }
                if (StartDate != null && StartDate != "" &&Endate!=null && Endate!="") {



               
                }
                else
                {
                    DateTime previousMonday = new DateTime();
                    DateTime lastSunday = DateTime.Now.AddDays(-1);
                    while (lastSunday.DayOfWeek != DayOfWeek.Sunday)
                    lastSunday = lastSunday.AddDays(-1);
                    previousMonday = lastSunday.AddDays(-6);
                    StartDate = previousMonday.ToString("yyyy-MM-dd");
                    Endate = lastSunday.ToString("yyyy-MM-dd");


                }
                AddVariable("StartDateofminibar", StartDate);
                AddVariable("endDateofminibar", Endate);
                Success("Bot executed sucessfully");



            }
            catch (Exception)

            {
                throw new Exception("Some Error Accoured...");
            }
        }
 
    }
}