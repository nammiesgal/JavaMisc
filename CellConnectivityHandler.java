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
import canrad.geometry.components.JunctionDevice;
import canrad.geometry.components.Port;
import canrad.geometry.components.Segment;
import canrad.geometry.components.Transceiver;
import canrad.layout.models.CellConnectionDetails;
import canrad.layout.models.FrequencyRange;
import canrad.layout.models.PortConnectionDetails;
import canrad.reference.components.AntennaTypePortSegment;
import canrad.reference.components.CanradLibrary;
import canrad.reference.components.Frequency;
import canrad.reference.components.MobilesCell;
import canrad.reference.components.MobilesFrequencyBand;
import canrad.reference.components.TransceiverFuncType;
import dialog.utilities.GlobalRoot;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Cheryl Long
 */
public class CellConnectivityHandler extends CellConnectivityHelper
{

    /**
     * For a given antenna port id, find all the transceiver ports connected to
     * it, and extract the Mobiles cells on
     * those ports
     * 
     * NOTE: This method should NOT be used as there is a bug with it! - It
     * doesn't retrieve all mobile cells/technologies that pass through the 
     * specified Antenna. It assumes a top down approach (Antenna down to 
     * Transceiver). Please use port.getCellConnectionDetails() - this code 
     * works on a bottom up approach (Transceiver up to Antenna) and will 
     * process ALL transceivers in the site export file and will filter the 
     * results based on the specified Antenna. 
     *
     * @param antenna       the antenna that contains the port
     * @param antennaPortId the port id to be traced
     * @return the list of mobile cells connected, will be empty if not
     *         connected
     */
    public static List<CellConnectionDetails> traceMobilesCellList(Antenna antenna, String antennaPortId)
    {
        try
        {
            // First, get all the connections for the antenna port and pull out the connection for the specified port
            PortConnectionDetails portConnection = getPortConnectionForDeviceAndPort(antenna, antennaPortId);
            if (portConnection == null)
            {
                return new ArrayList<>();
            }
            return getTransceiverPortsForAntennaPort(portConnection);
        }
        catch (Throwable t)
        {
            GlobalRoot.logError("CellConnectivityHandler.traceMobilesCellList", "Exception - Antenna: " + antenna.getId() + " - Port: " + antennaPortId, t);
            return new ArrayList<>();
        }
    }

    /**
     * Finds the connectivity details for a single antenna segment.
     * <p>
     * @param connection     Contains the connection details for the specific
     *                       antenna port. The connection passed in contains
     *                       the references to the antenna and the port, the feeder connected to it
     *                       and the device (junction device or transceiver)
     *                       on the other side.
     * @param antennaSegment The details about the segment being searched
     * @return The list of the cell connection details for the given segment
     */
    private static List<CellConnectionDetails> findConnectivityForSegment(PortConnectionDetails connection, Segment antennaSegment)
    {
        final CanradLibrary library = CanradLibrary.getInstance();
        final AntennaTypePortSegment typePortSegment = antennaSegment.getAntennaTypePortSegment();

        // Save the frequency range defined in the antenna port segment for use during the search
        FrequencyRange frequency = new FrequencyRange(typePortSegment.getMinFreq(), typePortSegment.getMaxFreq());
        List<FrequencyRange> frequencyRanges = new ArrayList<>();
        frequencyRanges.add(frequency);
        connection.getIncomingFrequencyRanges().clear();
        connection.getIncomingFrequencyRanges().addAll(frequencyRanges);

        // Traverse recursively through the connected devices to find the transceiver ports at the other end
        List<PortConnectionDetails> transceiverConnections = new ArrayList<>();
        List<Port> handledPorts = new ArrayList<>();
        getConnectedTransceiverPorts(connection, transceiverConnections, handledPorts, frequencyRanges);

        // If no transceivers were found, return the empty list
        if (transceiverConnections.isEmpty())
        {
            return new ArrayList<>();
        }

        // For each port / segment returned, ensure that there is a fit between the frequency passed through the
        // Junction Devices and the segment on the transceiver.  If there is a fit, save the cell details and connected
        // function details into the cell connection details list
        List<CellConnectionDetails> cellsList = new ArrayList<>();
        for (PortConnectionDetails trxConn : transceiverConnections)
        {
            Transceiver transceiver = ((Transceiver) trxConn.getConnectedDevice());
            if (!transceiver.isMobilesType())
            {
                continue;
            }

            Port trxPort = trxConn.getConnectedPort();
            for (Segment transceiverSegment : trxPort.getSegmentsList())
            {
                final Integer mobilesFreqBandId = transceiverSegment.getMobilesFreqBandId();
                final TransceiverFuncType funcType = library.getTransceiverFuncTypeById(transceiverSegment.getSegmentFuncTypeId());

                // If it has a mobiles frequency band, look up the frequency band and ensure that the antenna is in
                // the same frequency range
                Double minFrequency = null, maxFrequency = null;
                if (mobilesFreqBandId != null)
                {
                    final MobilesFrequencyBand frequencyBand = library.getMobilesFrequencyBandById(mobilesFreqBandId);
                    if (frequencyBand != null)
                    {
                        minFrequency = frequencyBand.getMinFrequency();
                        maxFrequency = frequencyBand.getMaxFrequency();
                    }
                }
                else
                {
                    final Frequency trxFrequency = library.getFrequency(transceiver.getFrequencyTypeId());
                    if (trxFrequency != null)
                    {
                        minFrequency = trxFrequency.getFreqMHz();
                        maxFrequency = trxFrequency.getPartnerMHz();
                    }
                }

                // If the frequency is in the correct range, include the details for the segment
                if (minFrequency != null && maxFrequency != null)
                {
                    if (isFrequencyFit(trxConn.getIncomingFrequencyRanges(), minFrequency, maxFrequency))
                    {
                        // Set up the cell connection details with the connected function type
                        CellConnectionDetails details = new CellConnectionDetails();
                        details.setConnectedPort(transceiverSegment.getParentPort());
                        details.setConnectedSegment(transceiverSegment);
                        details.setTransceiverFuncType(funcType);
                        if (transceiverSegment.getMobilesCellId() != null)
                        {
                            final MobilesCell mobilesCell = library.getMobilesCell(transceiverSegment.getMobilesCellId());
                            details.setMobilesCell(mobilesCell);
                        }
                        cellsList.add(details);
                    }
                }    
            }
        }
        return cellsList;
    }

    /**
     * Get all transceiver ports for the given feeder and recursively follow any
     * junction devices.
     *
     * @param feeder
     * @param segments
     */
    private static void getConnectedTransceiverPorts(PortConnectionDetails connection, List<PortConnectionDetails> transceiverConnections, List<Port> handledPorts, List<FrequencyRange> frequencyRanges)
    {

        // Keep track of the ports we have followed
        Port port = connection.getConnectedPort();
        if (handledPorts.contains(port))
        {
            return;
        }
        handledPorts.add(port);

        // If the connected device is a Transceiver, add it to the list and return
        if (connection.getConnectedDevice() instanceof Transceiver)
        {
            transceiverConnections.add(connection);
            return;
        }

        // If it is a junction device, find the mapped ports and call this recursively checking for a fit in frequency
        // of the device path traversed and the allowed frequency of the current junction device
        if (connection.getConnectedDevice() instanceof JunctionDevice)
        {
            JunctionDevice jd = (JunctionDevice) connection.getConnectedDevice();

            List<PortConnectionDetails> mappedPorts = findMappedPorts(jd, port, frequencyRanges);
            if (!mappedPorts.isEmpty())
            {
                mappedPorts.stream().forEach((nextConnection) ->
                {
                    if (nextConnection.getConnectedDevice() != null)
                    {
                        getConnectedTransceiverPorts(nextConnection, transceiverConnections, handledPorts, nextConnection.getIncomingFrequencyRanges());
                    }
                });
            }
        }
    }

    /**
     * For each segment of an antenna, search for and retrieve all transceiver
     * ports that are connected to the given
     * antenna port. Handles multiple jumps through any number of junction
     * devices along the way.
     */
    private static List<CellConnectionDetails> getTransceiverPortsForAntennaPort(PortConnectionDetails connection)
    {
        List<CellConnectionDetails> cellsList = new ArrayList<>();
        
        // Get the feeder_id, antenna_port_segment_id, segment_type.minimum_frequency, segment_type.maximum_frequency
        // for each connection to the port
        final Port port = connection.getPort();
        port.getSegmentsList().stream().forEach(antennaSegment ->
        {
            List<CellConnectionDetails> details = findConnectivityForSegment(connection, antennaSegment);
            cellsList.addAll(details);
        });
        return cellsList;
    }
}
