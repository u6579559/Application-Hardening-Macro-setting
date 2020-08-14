#!/bin/bash

#How the script should work.
#Take an argument stating whether you want all GPO Reports or a number of GPO reports
#First argument will be either "all" or "some"
#Following arguments will be the GPO's that will be reported (Only if some is the first argument)
#Example 1... ./GettingGPO.sh all
#Example 2... ./GettingGPO.sh some "DefaultGroupPolicy"
argc=$#
argv=("$@")

if [[ $# -eq 0 ]]
  then
    echo "No arguments supplied"
    exit 1
fi

if [[ $1 == "all" ]]
  then
    #Gets all GPO reports
    echo "Getting all GPOReports"
    Get-GPOReport -All -ReportType XML -Path AllGPOsReport.xml
elif [[ $1 == "some" ]]
  then
    echo "Getting some GPOReports"
    for (( i=1; i<argc; i++ )); do
      echo "Getting report for ${argv[i]}"
      Get-GPOReport -Name "${argv[i]}" -ReportType XML -Path ${argv[i]}Report.xml
    done
else
    echo "Incorrect command structure!"
fi