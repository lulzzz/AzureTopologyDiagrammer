# AzureTopologyDiagrammer
A Powershell-based Visio diagrammer for Azure environments. Currently tested in Windows 10 and Visio 2016. The goal of opening up this project is to encourage others to help build this into a great community solution, as I've been too hampered by time to get it to where it needs to be.

This initial version diagrams Azure regions and resource groups (with networking and subnets coming shortly). Connection to Azure is done via ADAL libraries and REST API instead of the Azure PowerShell module to provide increased flexibility.
