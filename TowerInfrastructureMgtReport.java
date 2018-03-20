/*
 * Copyright 2016 Telstra Corporation.
 *
 * This program is the property of Telstra Corporation.
 * You may not use this code without the express permission of
 * the designated business owner of the software in Telstra.
 * Parts of this program are copyright Dialog Information Technology
 * and licensed for the use of Telstra Corporation as part of this
 * software.
 */
package canrad.misc;

import canrad.geometry.components.Antenna;
import canrad.geometry.components.BatteryBank;
import canrad.geometry.components.DeviceBase;
import canrad.geometry.components.Feeder;
import canrad.geometry.components.GuyWire;
import canrad.geometry.components.Mount;
import canrad.geometry.components.PowerConverter;
import canrad.geometry.components.PowerSource;
import canrad.geometry.components.SiteExport;
import canrad.geometry.components.Structure;
import canrad.geometry.components.StructureBase;
import dialog.form.utils.FormHelper;
import dialog.utilities.GlobalRoot;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author SeetoB
 */
public class TowerInfrastructureMgtReport
{

    private final SiteExport siteExport;
    private final String fileName;
    private TIMReportGenerator timAuditReport;
    private FileOutputStream fileOut;

    public TowerInfrastructureMgtReport(SiteExport siteExport, String file)
    {
        this.siteExport = siteExport;
        fileName = file;
    }

    public void generateTIMAuditReport()
    {
        try
        {
            fileOut = new FileOutputStream(fileName);

            timAuditReport = new TIMReportGenerator();
            timAuditReport.setupReportStyles();
            
            List<String> dataList = new ArrayList<>();
            dataList.add(siteExport.getSite().getName());
            dataList.add(siteExport.getSite().getAddressId());
            if (siteExport.getSite().getCheckoutTime() == null)
            {
                if (siteExport.getSite().getExportTime() == null)
                    dataList.add("");
                else
                    dataList.add(siteExport.getSite().getExportTime().toString());
            }
            else
                dataList.add(siteExport.getSite().getCheckoutTime().toString());
            dataList.add(siteExport.getSite().getDatum());
            timAuditReport.createMainHeader(dataList);
            
            createReportBody();

            timAuditReport.autoSizeColumns();
            timAuditReport.writeToSpreadsheet(fileOut);
            fileOut.close();

        } catch (FileNotFoundException ex)
        {
            GlobalRoot.logError("InfrastructureManagementReport", "Streams - Exception: " + ex.getMessage(), ex);
            FormHelper.issueWarningAndWait(null, "Infrastructure Management Report Error",
                    "Problem creating Infrastructure Management Report: ",
                    "Please check specified file");
        } catch (IOException ex)
        {
            GlobalRoot.logError("InfrastructureManagementReport", "Streams - Exception: " + ex.getMessage(), ex);
        }
    }

    private void createReportBody()
    {
        createSiteCoords();
        createStructure();
        createStructWires();
        createAntenna();
        createFeeder();
        createShelter();
        createTransceiver();
        createBatteryBank();
        createPowerConverter();
        createPrimaryPowerSource();
    }

    private void createSiteCoords()
    {
        timAuditReport.createSiteCoordDetailHdr();
        List<String> dataList = new ArrayList<>();
        String coords = "Lat: " + siteExport.getSite().getLatitude() + " Long: " + siteExport.getSite().getLongitude();       
        String zone = String.valueOf(Coords.spheroidToZone(siteExport.getSite().getLongitude(), siteExport.getSite().getDatum()));
        String easting = String.valueOf(Coords.spheroidToEasting(siteExport.getSite().getLatitude(), siteExport.getSite().getLongitude(), siteExport.getSite().getDatum()));
        String northing = String.valueOf(Coords.spheroidToNorthing(siteExport.getSite().getLatitude(), siteExport.getSite().getLongitude(), siteExport.getSite().getDatum()));
        dataList.add(coords);
        dataList.add(siteExport.getSite().getDatum());
        dataList.add(siteExport.getSite().getCoordinatePrecision());
        dataList.add(zone);
        dataList.add(easting);
        dataList.add(northing);
        dataList.add(siteExport.getSite().getRelativeLevel().toString());
        timAuditReport.createSiteCoordDetailRow(dataList);
    }

    private void createStructure()
    {
        timAuditReport.createStructureDetailHdr();

        for (StructureBase structureBase : siteExport.getStructuresNoAddRemove())
        {
            if (!Structure.class.isAssignableFrom(structureBase.getClass()))
                continue;
            Structure structure = (Structure) structureBase;
            List<String> dataList = new ArrayList<>();
            dataList.add(structure.getNumber());
            dataList.add(structure.getTypeCode());
            dataList.add(String.valueOf(structure.getHeight()));
            String extHeightStr = "";
            if (structure.getExtHeight() != null)
            {
                extHeightStr = " (" + structure.getExtHeight() + "m)";
            }
            dataList.add(structure.getExtType() + extHeightStr);
            dataList.add(structure.getPainted());
            dataList.add(structure.getWarningLights());
            dataList.add(structure.getAccessRestriction());
            dataList.add(structure.getOwnerCode());
            dataList.add(structure.getStatusCode());
            dataList.add(structure.getNotes());

            timAuditReport.createStructureDetailRow(dataList);
        }
    }

    private void createStructWires()
    {
        timAuditReport.createStructureGuyWiresDetailHdr();
        boolean found = false;
        for (StructureBase structureBase : siteExport.getStructuresNoAddRemove())
        {
            if (!Structure.class.isAssignableFrom(structureBase.getClass()))
                continue;
            Structure structure = (Structure) structureBase;
            for(GuyWire wire : structure.getGuyWires())
            {
                found = true;
                List<String> dataList = new ArrayList<>();
                dataList.add(structure.getNumber());
                dataList.add(wire.getCableNumber());
                dataList.add(String.valueOf(wire.getAttachHeight()));
                dataList.add(wire.getFaceOrCorner());
                dataList.add(wire.getStatusCode());
                dataList.add(String.valueOf(wire.getBearing()));
                dataList.add(String.valueOf(wire.getKNDesigned()));
                dataList.add(String.valueOf(wire.getKNMeasured()));
                dataList.add(wire.getCableType());
                dataList.add("");
                dataList.add(wire.getTorqueStabilised());
                timAuditReport.createStructureGuyWiresDetailRow(dataList);
            }
        }
        if (!found)
        {
            List<String> dataList = new ArrayList<>();
            timAuditReport.createStructureGuyWiresDetailRow(dataList);
        }
        //TODO Populate data here!
        
    }

    private void createAntenna()
    {
        timAuditReport.createAntennaDetailHdr();
        for (DeviceBase antenna : siteExport.getAntennasNoAddRemove())
        {
            List<String> dataList = new ArrayList<>();
            dataList.add(antenna.getNumber());
            String attachType = "none";
            if (((Antenna) antenna).getAttachType() != null)
            {
                String tempStr = ((Antenna) antenna).getAttachType().replaceAll("\\s", "");
                if (!tempStr.isEmpty())
                {
                    attachType = ((Antenna) antenna).getAttachType();
                }
            }
            dataList.add(attachType);
            dataList.add(antenna.getTypeCode());
            dataList.add(antenna.getBearing().toString());
            dataList.add(((Antenna) antenna).getPolarizationTypeCode());
            dataList.add(antenna.getFace());
            String height;
            if (((Antenna) antenna).getMount() != null)
            {
                String mountPos = ((Antenna) antenna).getMountPos();
                Mount mount = ((Antenna) antenna).getMount();
                height = String.valueOf(mount.getMountPoint(mountPos).getHeightOnStructure());
            }
            else
            {
                Double val = antenna.getHeight();
                height = (val == null) ? "Not Specified" : val.toString();
            }
            String structNum = "";
            if (antenna.getStructure() != null)
            {
                structNum = antenna.getStructure().getNumber();
            }
            else if (antenna.getMount() != null)
            {
                Mount mount = antenna.getMount();
                structNum = mount.getStructure().getNumber() + " / " + mount.getNumber() + " / " + antenna.getMountPos();
            }
            dataList.add(height);
            String midPoint;
            if (((Antenna)antenna).getAntennaMidPoint() == null)
                midPoint = "Not Specified";
            else
                midPoint = String.valueOf(((Antenna)antenna).getAntennaMidPoint());
            dataList.add(midPoint);
            dataList.add(structNum);
            dataList.add(antenna.getOwnerCode());
            dataList.add(antenna.getStatusCode());

            timAuditReport.createAntennaDetailRow(dataList);
        }
    }

    private void createFeeder()
    {
        timAuditReport.createFeederDetailHdr();
        for (Feeder feeder : siteExport.getFeedersNoAddRemove())
        {
            List<String> dataList = new ArrayList<>();
            dataList.add(feeder.getNumber());
            dataList.add(feeder.getTypeCode());
            dataList.add(feeder.getLength().toString());
            dataList.add(feeder.getOwnerCode());
            DeviceBase device = feeder.getFeederAEnd().getDevice();
            String val;
            if (device != null)
                val = device.getNumber();
            else
                val = "";
            dataList.add(val);
            dataList.add("");
            device = feeder.getFeederBEnd().getDevice();
            if (device != null)
                val = device.getNotes();
            else
                val = "";
            dataList.add(val);
            dataList.add("");
            dataList.add(feeder.getStatusCode());
            dataList.add(feeder.getNotes());

            timAuditReport.createFeederDetailRow(dataList);
        }
    }

    private void createShelter()
    {
        timAuditReport.createShelterDetailHdr();
        for (StructureBase shelter : siteExport.getSheltersNoAddRemove())
        {
            List<String> dataList = new ArrayList<>();
            dataList.add(shelter.getNumber());
            dataList.add(shelter.getTypeCode());
            dataList.add("");
            dataList.add("");
            dataList.add(shelter.getStatusCode());
            dataList.add(shelter.getOwnerCode());
            dataList.add("");
            dataList.add(shelter.getNotes());

            timAuditReport.createShelterDetailRow(dataList);
        }
    }

    private void createTransceiver()
    {
        timAuditReport.createTransceiverDetailHdr();
        for (DeviceBase transceiver : siteExport.getTransceiversNoAddRemove())
        {
            List<String> dataList = new ArrayList<>();
            dataList.add(transceiver.getNumber());
            dataList.add(transceiver.getTypeCode());
            dataList.add("");
            dataList.add("");
            dataList.add(transceiver.getStatusCode());
            dataList.add(transceiver.getOwnerCode());
            dataList.add("");
            dataList.add(transceiver.getNotes());

            timAuditReport.createTransDetailRow(dataList);
        }
    }

    private void createBatteryBank()
    {
        boolean found = false;
        timAuditReport.createBatteryBankDetailHdr();
        
        for(BatteryBank bank : siteExport.getSite().getBatteryBanks())
        {
            found = true;
            List<String> dataList = new ArrayList<>();
            dataList.add(bank.getBankNumber());
            dataList.add(bank.getBankType());
            dataList.add("");
            dataList.add(bank.getBatteryCount().toString());
            dataList.add(String.valueOf(bank.getVoltage()));
            dataList.add("");
            dataList.add("");
            dataList.add(bank.getShelterNumber());
            dataList.add(bank.getStatusCode());
            dataList.add(bank.getNotes());
            timAuditReport.createBatteryBankDetailRow(dataList);
        }
        
        if (!found)
        {
            List<String> dataList = new ArrayList<>();
            timAuditReport.createBatteryBankDetailRow(dataList);
        }
    }

    private void createPowerConverter()
    {
        boolean found = false;
        timAuditReport.createPowerConverterDetailHdr();
        for(PowerConverter pc : siteExport.getSite().getPowerConverters())
        {
            found = true;
            List<String> dataList = new ArrayList<>();
            dataList.add(pc.getConverterNumber());
            dataList.add(pc.getConverterType());
            dataList.add("");
            dataList.add("");
            dataList.add(pc.getStatusCode());
            dataList.add(pc.getShelterNumber());            
            dataList.add(pc.getNotes());            
            timAuditReport.createPowerConverterDetailRow(dataList);
        }
        
        if (!found)
        {
            List<String> dataList = new ArrayList<>();
            timAuditReport.createPowerConverterDetailRow(dataList);
        }
    }

    private void createPrimaryPowerSource()
    {
        boolean found = false;
        timAuditReport.createPrimaryPowerSourceDetailHdr();
        for(PowerSource ps : siteExport.getSite().getPowerSources())
        {
            found = true;
            List<String> dataList = new ArrayList<>();
            dataList.add(ps.getPowerSource());
            dataList.add("");
            dataList.add(ps.getPowerSourceType());
            dataList.add("");
            dataList.add(ps.getCount().toString());
            dataList.add("");
            dataList.add(String.valueOf(ps.getVoltage()));
            dataList.add(ps.getStatus());
            dataList.add(ps.getSerialNumbers());
            timAuditReport.createPrimaryPowerSourceDetailRow(dataList);
        }
        
        if (!found)
        {
            List<String> dataList = new ArrayList<>();
            timAuditReport.createPrimaryPowerSourceDetailRow(dataList);
        }
    }
}