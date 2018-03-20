/*
 * Copyright 2015 Telstra Corporation.
 *
 * This program is the property of Telstra Corporation.
 * You may not use this code without the express permission of
 * the designated business owner of the software in Telstra.
 * Parts of this program are copyright Dialog Information Technology
 * and licensed for the use of Telstra Corporation as part of this
 * software.
 */
package canrad.celltrace;

import canrad.geometry.components.Antenna;
import canrad.geometry.components.DeviceBase;
import canrad.geometry.components.JunctionDevice;
import canrad.geometry.components.Port;
import canrad.geometry.components.SiteExport;
import canrad.geometry.components.SiteExportViewModel;
import canrad.geometry.components.Transceiver;
import canrad.layout.models.CellConnectionDetails;
import canrad.layout.models.FrequencyRange;
import canrad.layout.models.PortConnectionDetails;
import canrad.reference.components.CanradLibrary;
import canrad.reference.components.JunctionDeviceTypeIntPortMap;
import canrad.reference.components.MobilesCell;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 *
 * @author Cheryl Long
 */
public abstract class CellConnectivityHelper
{

    /**
     * Adjust the allowed frequency range
     * <p>
     * @param mapping
     * @param frequencyRanges
     * @return
     */
    protected static List<FrequencyRange> adjustAllowedFrequencyRange(JunctionDeviceTypeIntPortMap mapping, List<FrequencyRange> frequencyRanges)
    {
        List<FrequencyRange> remainingFrequencies = new ArrayList<>();
        boolean isPass = "PASS".equalsIgnoreCase(mapping.getPassOrStop());
        // If there are no frequencies on the mapping record and it is a pass, all existing frequencies carry on through.
        // If it is a stop, none within the given stop range shall pass.
        if (mapping.getMinFrequency() == null || mapping.getMaxFrequency() == null)
        {
            if (isPass)
            {
                remainingFrequencies.addAll(frequencyRanges);
            }
            return remainingFrequencies;
        }
        // If the mapping record has frequencies, calculate if there are any frequency ranges that are allowed to pass through.
        frequencyRanges.stream().forEach((FrequencyRange range) ->
        {
            calculateFrequencyRanges(mapping, range.getMinFrequency(), range.getMaxFrequency(), isPass, remainingFrequencies);
        });
        return remainingFrequencies;
    }

    /**
     * Takes the junction device mapping record, the incoming minimum and
     * maximum frequency range and the pass/stop flag
     * and determines which frequency ranges (if any) will pass through the
     * device.
     *
     * @param mapping              The junction device port mapping record
     * @param min                  The minimum incoming frequency
     * @param max                  The maximum incoming frequency
     * @param isPass               The pass condition flag
     * @param remainingFrequencies The list of frequencies that remain after the
     *                             mapping is applied
     */
    protected static void calculateFrequencyRanges(JunctionDeviceTypeIntPortMap mapping, Double min, Double max, Boolean isPass, List<FrequencyRange> remainingFrequencies)
    {
        // The current CANRAD system does not enforce this condition (min <= max) on the Antenna frequency range.
        // Given that this is an invalid condition and shouldn't happen, don't check further for a fit.
        if (min > max)
        {
            return;
        }
        if (isPass)
        {
            // A pass condition indicates that frequencies that are inside the range of the junction device will pass through
            // if there is any overlap between the incoming frequency and the mapping frequency.  The remaining frequency
            // will be the overlap area between the two.
            if (hasFrequencyOverlap(mapping, min, max))
            {
                Double newMin = mapping.getMinFrequency() > min ? mapping.getMinFrequency() : min;
                Double newMax = mapping.getMaxFrequency() < max ? mapping.getMaxFrequency() : max;
                FrequencyRange newFrequency = new FrequencyRange(newMin, newMax);
                remainingFrequencies.add(newFrequency);
            }
        }
        else
        {
            // Is stop condition.  This indicates that any frequencies that are inside the range of the mapping frequency
            // range cannot pass through.  The remaining frequencies will be the area outside the mapping range. This
            // can be split in two if the incoming range completely spans the mapping range.
            if (!isFrequencyFullyContained(mapping, min, max))
            {
                // If there is no overlap, the full incoming frequency can pass through
                if (hasNoOverlap(mapping, min, max))
                {
                    FrequencyRange range = new FrequencyRange(min, max);
                    remainingFrequencies.add(range);
                    return;
                }
                // The mapping frequency is completely contained in the incoming frequency, the remaining frequency is
                // split into the area above and below the mapping range
                if (min < mapping.getMinFrequency() && max > mapping.getMaxFrequency())
                {
                    FrequencyRange lowerRange = new FrequencyRange(min, mapping.getMinFrequency());
                    FrequencyRange upperRange = new FrequencyRange(mapping.getMaxFrequency(), max);
                    remainingFrequencies.add(lowerRange);
                    remainingFrequencies.add(upperRange);
                    return;
                }
                // There is a single area above or below the mapped frequency range
                FrequencyRange range;
                if (max > mapping.getMaxFrequency())
                {
                    range = new FrequencyRange(mapping.getMaxFrequency(), max);
                }
                else
                {
                    range = new FrequencyRange(min, mapping.getMinFrequency());
                }
                remainingFrequencies.add(range);
            }
        }
    }

    /**
     * For a given junction device and port, use the internal port mapping to
     * find the port or ports that come out
     * the other side of the input port. Uses the port mapping rules (PASS /
     * STOP) and the frequency range to determine
     * whether the signal will go out any of the ports or will stop there.
     *
     * If there is a frequency match, the port connection details will be passed
     * back with the revised frequency range(s)
     * that pass through this device.
     *
     * @param jd              The junction device being checked
     * @param port            The input port on the device
     * @param frequencyRanges The frequency ranges that are still available for
     *                        this path
     * @return The port connection details for all ports that are connected and
     *         have a frequency match
     */
    protected static List<PortConnectionDetails> findMappedPorts(JunctionDevice jd, Port port, List<FrequencyRange> frequencyRanges)
    {
        List<PortConnectionDetails> connectionMappingList = new ArrayList<>();
        
        final String deviceTypeCode = jd.getTypeCode();
        final String typePortId = port.getTypePortId();
        final List<JunctionDeviceTypeIntPortMap> portMappings = CanradLibrary.getInstance().getJDIntPortMaps(deviceTypeCode);
        Map<Port, PortConnectionDetails> connectionMap = new LinkedHashMap<>();
        for (JunctionDeviceTypeIntPortMap mapping : portMappings)
        {
            String otherPortTypeId = null;
            if (mapping.getPortA().equals(typePortId))
            {
                otherPortTypeId = mapping.getPortB();
            }
            else if (mapping.getPortB().equals(typePortId))
            {
                otherPortTypeId = mapping.getPortA();
            }
            // If there are frequencies on the port mapping, need to check the fit with the current frequency range(s)
            if (otherPortTypeId != null)
            {
                List<FrequencyRange> newFrequencyRanges = adjustAllowedFrequencyRange(mapping, frequencyRanges);
                if (newFrequencyRanges.isEmpty())
                {
                    continue;
                }
                // If we have found a port mapping and it either has no frequency defined or passes the frequency fit check...
                // If the port passes and the frequency passes,
                Port otherPort = jd.getPortByTypePortId(otherPortTypeId);
                PortConnectionDetails portConnection;
                if (connectionMap.containsKey(otherPort))
                {
                    portConnection = connectionMap.get(otherPort);
                }
                else
                {
                    portConnection = getPortConnectionForDeviceAndPort(jd, otherPort.getId());
                    portConnection.getIncomingFrequencyRanges().clear();
                }
                
                //TODO 
                portConnection.addMapping(mapping);
                portConnection.getIncomingFrequencyRanges().addAll(newFrequencyRanges);
                connectionMap.put(otherPort, portConnection);
            }
        }
        connectionMap.values().stream().forEach((PortConnectionDetails portConnection) ->
        {
            connectionMappingList.add(portConnection);
        });
        return connectionMappingList;
    }

    /**
     * Retrieve the connection details for the current port on a device.
     *
     * @param device The current device (antenna or junction device)
     * @param portId The id of the port on the device
     * @return The PortConnectionDetails for the requested device (null if not
     *         found). Structure contain details about
     *         the device and port, the feeder connected and the device and port on the
     *         other side.
     */
    protected static PortConnectionDetails getPortConnectionForDeviceAndPort(DeviceBase device, String portId)
    {
        List<PortConnectionDetails> connections = device.buildPortConnectionsList();
        for (PortConnectionDetails connection : connections)
        {
            if (connection.getPort().getId().equals(portId))
            {
                return connection;
            }
        }
        return null;
    }

    /**
     * Check if the incoming frequency range has an overlap with the mapping
     * frequency range.
     *
     * @param targetMin    the min frequency of the mapping
     * @param targetMax    the max frequency of the mapping
     * @param minFrequency the minimum frequency of the incoming device
     * @param maxFrequency the maximum frequency of the incoming device
     * @return true if there is an overlap between the frequency range in the
     *         mapping record and the incoming frequency range
     */
    protected static boolean hasFrequencyOverlap(Double targetMin, Double targetMax, Double minFrequency, Double maxFrequency)
    {
        return minFrequency <= targetMax && maxFrequency >= targetMin;
    }

    /**
     * Check if the incoming frequency range has an overlap with the mapping
     * frequency range.
     *
     * @param mapping      the junction device internal port map record. The min
     *                     and max frequencies must not be null
     * @param minFrequency the minimum frequency of the incoming device
     * @param maxFrequency the maximum frequency of the incoming device
     * @return true if there is an overlap between the frequency range in the
     *         mapping record and the incoming frequency range
     */
    protected static boolean hasFrequencyOverlap(JunctionDeviceTypeIntPortMap mapping, Double minFrequency, Double maxFrequency)
    {
        return hasFrequencyOverlap(mapping.getMinFrequency(), mapping.getMaxFrequency(), minFrequency, maxFrequency);
    }

    /**
     * Check if the incoming frequency range is completely outside the range of
     * the mapping frequency range.
     *
     * @param mapping      the junction device internal port map record. The min
     *                     and max frequencies must not be null
     * @param minFrequency the minimum frequency of the incoming device
     * @param maxFrequency the maximum frequency of the incoming device
     * @return true if there is an overlap between the frequency range in the
     *         mapping record and the incoming frequency range
     */
    protected static boolean hasNoOverlap(JunctionDeviceTypeIntPortMap mapping, Double minFrequency, Double maxFrequency)
    {
        return minFrequency > mapping.getMaxFrequency() || maxFrequency < mapping.getMinFrequency();
    }

    /**
     * Calculates whether there is a "fit" between the frequency ranges that
     * have arrived at this device and the min / max
     * allowed on the receiving device. There is a fit if there is any overlap
     * between the ranges.
     *
     * @param frequencyRanges The list of frequency ranges that have arrived at
     *                        this device
     * @param min             The minimum frequency that this device allows
     * @param max             The maximum frequency that this device allows
     * @return true if there is an overlap between any of the frequency ranges
     *         provided and the min / max on this device
     */
    protected static boolean isFrequencyFit(List<FrequencyRange> frequencyRanges, Double min, Double max)
    {
        for (FrequencyRange range : frequencyRanges)
        {
            if (hasFrequencyOverlap(range.getMinFrequency(), range.getMaxFrequency(), min, max))
            {
                return true;
            }
        }
        return false;
    }

    /**
     * Check if the incoming frequency range is fully contained within the
     * mapping frequency range.
     *
     * @param mapping      the junction device internal port map record. The min
     *                     and max frequencies must not be null
     * @param minFrequency the minimum frequency of the incoming device
     * @param maxFrequency the maximum frequency of the incoming device
     * @return true if there is an overlap between the frequency range in the
     *         mapping record and the incoming frequency range
     */
    protected static boolean isFrequencyFullyContained(JunctionDeviceTypeIntPortMap mapping, Double minFrequency, Double maxFrequency)
    {
        return minFrequency >= mapping.getMinFrequency() && maxFrequency <= mapping.getMaxFrequency();
    }
    
    public static String getAntennaCellInfo(Antenna antenna)
    {
        List<Port> portList = new ArrayList();
        List<String> mobileCellList = new ArrayList();
        List<CellConnectionDetails> traceMobilesCellList = antenna.getCellConnectionDetails();       
        for (CellConnectionDetails cellConnectionDetails : traceMobilesCellList)
        {
            if (cellConnectionDetails.getMobilesCell() != null && !mobileCellList.contains(cellConnectionDetails.getMobilesCell().getName()))
                mobileCellList.add(cellConnectionDetails.getMobilesCell().getName());
        }
        
        if (!mobileCellList.isEmpty())
        {    
            String allMobileCells = "(Cells : ";
            for (String s : mobileCellList)
            {
                allMobileCells = allMobileCells.concat(s + ";");
            }
            allMobileCells = allMobileCells.substring(0, allMobileCells.length()-1);
            allMobileCells = allMobileCells.concat(")");
            return allMobileCells;
        }
        
        return null;
    }  
}
