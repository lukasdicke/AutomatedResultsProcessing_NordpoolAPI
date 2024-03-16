# title: "AutomatedResultsProcessingNordpoolAPI"
# description: "This script executes the NordpoolAPI console application for results download and export to PIXOS (e.g. 'nordpool_nl')."
# output: ""
# parameters: {}
# owner: "MECO, Lukas Dicke"

"""

Usage:
    AutomatedResultsProcessingNordpoolAPI.py <job_path> --exchangeName=<string>

Options:
    --exchangeName=<string> Matchname of exchange name, see file 'ConfigDataNordpool.xml' in known location.

"""

import subprocess
import sys

#exchangeName = "nordpool_nl"

exchangeName = str(exchangeName)

uncPath = "\\\\energycorp.com\\common\\Divsede\\Operations\\Personal_OPS\\Lukas\\DevelopedApplications\\NordpoolAPI_CWE\\ResultProcessing_NordpoolApi_CWE\\bin\\Debug\\ResultProcessing_NordpoolApi_CWE.exe"

process = subprocess.run(uncPath + " " + exchangeName, stdout=True)



if process.returncode != 0 :
    print("Something failed, when downloading the exchange-result: '" + exchangeName+"' (see Return-code below)")
    print("Return-code: " + str( process.returncode))
    sys.exit(1)
else:
    print("Successful result download!")