"""
Copyright (c) 2018 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.0 (the "License"). You may obtain a copy of the
License at
               https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""
from jinja2 import Environment
from jinja2 import FileSystemLoader
import os
import requests
import json
import xlsxwriter
import base64

DIR_PATH = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
JSON_TEMPLATES = Environment(loader=FileSystemLoader(DIR_PATH + '/json_templates'))

# Disable https warnings
requests.packages.urllib3.disable_warnings()


class PrimeController:
    url = ""
    username = ""
    password = ""

    def makeCall(self, p_url, method, data=""):
        """
        Single exit point for all APIs calls for Prime Infrastructure
        :param p_url:
        :param method:
        :param data:
        :return:
        """
        credentials = self.username + ":" + self.password

        headers = {
            'Authorization': "Basic " + base64.b64encode(bytes(credentials, "utf-8")).decode("utf-8")
        }
        if method == "POST":
            response = requests.post(self.url + p_url, data=data, headers=headers, verify=False)
        elif method == "GET":
            response = requests.get(self.url + p_url, headers=headers, verify=False)
        else:
            raise Exception("Method " + method + " not supported by this controller")
        if 199 > response.status_code > 300:
            errorMessage = json.loads(response.text)["errorDocument"]["message"]
            raise Exception("Error: status code" + str(response.status_code) + " - " + errorMessage)
        return response

    def getAPs(self):
        """
        Return a list of access points
        :return:
        """
        pURL = "/webacs/api/v3/data/AccessPointDetails.json"
        response = self.makeCall(p_url=pURL, method="GET")
        print(response.text)
        APs = json.loads(response.text)["queryResponse"]["entityId"]
        return APs

    def getAPDetail(self, apDetailId):
        """
        Return detail data of an specific access point
        :param apDetailId:
        :return:
        """
        pURL = "/webacs/api/v3/data/AccessPointDetails/" + apDetailId + ".json"
        response = json.loads(self.makeCall(p_url=pURL, method="GET").text)
        APDetail = None
        if len(response["queryResponse"]["entity"]) > 0:
            APDetail = json.loads(self.makeCall(p_url=pURL, method="GET").text)["queryResponse"]["entity"][0]
        return APDetail

    def getClientCount(self):
        """
        Returns the sum of all 2.4G and 5G clients
        :return:
        """
        fiveGCount = 0
        twoPointFourGCount = 0
        # Get APs
        APs = self.getAPs()
        for AP in APs:
            apDetail = self.getAPDetail(apDetailId=AP["$"])
            print(apDetail)
            fiveGCount += int(apDetail["accessPointDetailsDTO"]["clientCount_5GHz"])
            twoPointFourGCount += int(apDetail["accessPointDetailsDTO"]["clientCount_2_4GHz"])
            # Return details
        return {
            "fiveGClients": fiveGCount,
            "twoPointFourGClients": twoPointFourGCount
        }

    def getRFLoadStats(self):
        """
         Return the RF Load stats list. Mainly for ChannelUtilization and PoorCoverageClients
        """
        result = []
        pURL = "/webacs/api/v3/data/RFLoadStats.json"
        rfLoadStatsList = json.loads(self.makeCall(p_url=pURL, method="GET").text)["queryResponse"]["entityId"]
        for rfLoadStat in rfLoadStatsList:
            pURL = rfLoadStat["@url"].replace(self.url, "") + ".json"
            response = json.loads(self.makeCall(p_url=pURL, method="GET").text)
            if len(response) > 0:
                rfLoadDetail = response["queryResponse"]["entity"][0]
                result.append(rfLoadDetail)
        return result

    def getRFStats(self):
        """
        Gets RF Stats data. Mainly for PowerOutput
        :return:
        """
        result = []
        pURL = "/webacs/api/v3/data/RFStats.json"
        rfStatsList = json.loads(self.makeCall(p_url=pURL, method="GET").text)["queryResponse"]["entityId"]
        for rfStat in rfStatsList:
            pURL = rfStat["@url"].replace(self.url, "") + ".json"
            response = json.loads(self.makeCall(p_url=pURL, method="GET").text)
            if len(response) > 0:
                rfStatDetail = response["queryResponse"]["entity"][0]
                result.append(rfStatDetail)
        return result

    def getRFCounters(self):
        """
        Gets RF Counters data. Mainly for multipleRetryCount, retryCount, rxFragmentCount and txFragmentCount
        :return:
        """
        result = []
        pURL = "/webacs/api/v3/data/RFCounters.json"
        rfCountersList = json.loads(self.makeCall(p_url=pURL, method="GET").text)["queryResponse"]["entityId"]
        for rfCounter in rfCountersList:
            pURL = rfCounter["@url"].replace(self.url, "") + ".json"
            response = json.loads(self.makeCall(p_url=pURL, method="GET").text)
            if len(response) > 0:
                rfCounterDetail = response["queryResponse"]["entity"][0]
                result.append(rfCounterDetail)
        return result

    def getWirelessClientSessions(self):
        result = []
        pURL = "/webacs/api/v3/data/ClientSessions.json"
        sessions = json.loads(self.makeCall(p_url=pURL, method="GET").text)["queryResponse"]["entityId"]
        for session in sessions:
            pURL = session["@url"].replace(self.url, "") + ".json"
            response = json.loads(self.makeCall(p_url=pURL, method="GET").text)
            if len(response) > 0:
                sessionDetail = response["queryResponse"]["entity"][0]
                if "apMacAddress" in sessionDetail["clientSessionsDTO"].keys():
                    result.append(sessionDetail)
        return result

    def startCollection(self):
        """
        Collects data from Prime and save it to a database
        :return:
        """

        workbook = xlsxwriter.Workbook(
            'prime-kpis-' + self.url.replace("http://", '').replace("https://", "").replace("/", "") + '.xlsx')
        worksheet = workbook.add_worksheet()

        row = 0
        col = 0

        worksheet.write(row, col, "AP Name")
        worksheet.write(row, col + 1, "5 GHz clients")
        worksheet.write(row, col + 2, "2 GHz clients")
        worksheet.write(row, col + 3, "Channel utilization")
        worksheet.write(row, col + 4, "Poor coverage clients")
        worksheet.write(row, col + 5, "Tx power output")
        worksheet.write(row, col + 6, "Tx fragment count")
        worksheet.write(row, col + 7, "Rx fragment count")
        worksheet.write(row, col + 8, "Retry count")
        worksheet.write(row, col + 9, "Multiple retry count")
        worksheet.write(row, col + 10, "Total bytes received")
        worksheet.write(row, col + 11, "Total bytes sent")

        row += 1

        print("Starting collection... ")
        # Access Points
        APs = self.getAPs()
        print("Found " + str(len(APs)) + " access points")
        print("Getting the RF Stats")
        RFLoadStatsList = self.getRFStats()
        print("Getting the RF Counters")
        RFCountList = self.getRFCounters()
        print("Getting the RF Load Stats")
        RFLoadDetailList = self.getRFLoadStats()
        print("Getting client sessions")
        clientSessions = self.getWirelessClientSessions()

        progress = 0
        print("Adding collection to database... ")
        print("Progress 0%")
        for AP in APs:

            apDetail = self.getAPDetail(apDetailId=AP["$"])
            totalBytesReceived = 0
            totalBytesSent = 0

            AP["rfLoads"] = []
            AP["rfStats"] = []
            AP["rfCounts"] = []

            for RFLoadDetail in RFLoadDetailList:
                if RFLoadDetail["rfLoadStatsDTO"]["macAddress"] == apDetail["accessPointDetailsDTO"]["macAddress"]:
                    ChannelUtilization = RFLoadDetail["rfLoadStatsDTO"]["channelUtilization"]
                    PoorCoverageClients = RFLoadDetail["rfLoadStatsDTO"]["poorCoverageClients"]
                    slotId = RFLoadDetail["rfLoadStatsDTO"]["slotId"]
                    AP["rfLoads"].append(
                        {
                            "ChannelUtilization": str(ChannelUtilization),
                            "PoorCoverageClients": str(PoorCoverageClients),
                            "slotId": str(slotId)
                        })

            for RFStatsDetail in RFLoadStatsList:
                if RFStatsDetail["rfStatsV3DTO"]["macAddress"] == apDetail["accessPointDetailsDTO"]["macAddress"]:
                    txPowerOutput = RFStatsDetail["rfStatsV3DTO"]["txPowerOutput"]
                    channelNumber = RFStatsDetail["rfStatsV3DTO"]["channelNumber"]
                    slotId = RFStatsDetail["rfStatsV3DTO"]["slotId"]
                    AP["rfStats"].append(
                        {
                            "txPowerOutput": str(txPowerOutput),
                            "channelNumber": str(channelNumber),
                            "slotId": str(slotId)
                        })

            for RFCount in RFCountList:
                if RFCount["rfCountersDTO"]["macAddress"] == apDetail["accessPointDetailsDTO"]["macAddress"]:
                    txFragmentCount = RFCount["rfCountersDTO"]["txFragmentCount"]
                    rxFragmentCount = RFCount["rfCountersDTO"]["rxFragmentCount"]
                    retryCount = RFCount["rfCountersDTO"]["retryCount"]
                    multipleRetryCount = RFCount["rfCountersDTO"]["multipleRetryCount"]
                    slotId = RFCount["rfCountersDTO"]["slotId"]
                    AP["rfCounts"].append(
                        {
                            "txFragmentCount": str(txFragmentCount),
                            "rxFragmentCount": str(rxFragmentCount),
                            "retryCount": str(retryCount),
                            "multipleRetryCount": str(multipleRetryCount),
                            "slotId": str(slotId)
                        })

            for session in clientSessions:
                if session["clientSessionsDTO"]["apMacAddress"] == apDetail["accessPointDetailsDTO"]["macAddress"]:
                    totalBytesReceived += int(session["clientSessionsDTO"]["bytesReceived"])
                    totalBytesSent += int(session["clientSessionsDTO"]["bytesSent"])

            ChannelUtilizationStr = ""
            PoorCoverageClientsStr = ""
            for item in AP["rfLoads"]:
                ChannelUtilizationStr += "Slot " + item["slotId"] + ": " + item["ChannelUtilization"] + " - "
                PoorCoverageClientsStr += "Slot " + item["slotId"] + ": " + item["PoorCoverageClients"] + " - "

            txPowerOutputStr = ""
            for item in AP["rfStats"]:
                txPowerOutputStr += "Slot " + item["slotId"] + ": Channel " + item["channelNumber"] + " " + item[
                    "txPowerOutput"] + " - "

            txFragmentCountStr=""
            rxFragmentCountStr=""
            retryCountStr=""
            multipleRetryCountStr=""
            for item in AP["rfCounts"]:
                txFragmentCountStr += "Slot " + item["slotId"] + ": " + item["txFragmentCount"] + " - "
                rxFragmentCountStr += "Slot " + item["slotId"] + ": " + item["rxFragmentCount"] + " - "
                retryCountStr += "Slot " + item["slotId"] + ": " + item["retryCount"] + " - "
                multipleRetryCountStr += "Slot " + item["slotId"] + ": " + item["multipleRetryCount"] + " - "

            worksheet.write(row, col, apDetail["accessPointDetailsDTO"]["name"])
            worksheet.write(row, col + 1, apDetail["accessPointDetailsDTO"]["clientCount_5GHz"])
            worksheet.write(row, col + 2, apDetail["accessPointDetailsDTO"]["clientCount_2_4GHz"])
            worksheet.write(row, col + 3, ChannelUtilizationStr)
            worksheet.write(row, col + 4, PoorCoverageClientsStr)
            worksheet.write(row, col + 5, txPowerOutputStr)
            worksheet.write(row, col + 6, txFragmentCountStr)
            worksheet.write(row, col + 7, rxFragmentCountStr)
            worksheet.write(row, col + 8, retryCountStr)
            worksheet.write(row, col + 9, multipleRetryCountStr)
            worksheet.write(row, col + 10, totalBytesReceived)
            worksheet.write(row, col + 11, totalBytesSent)
            row += 1

            progress += 1
            print("Progress: " + "%.2f" % (progress / len(APs) * 100) + "%")
        worksheet.autofilter('A1:L' + str(row + 2))
        worksheet.set_column("A:L", 20)
