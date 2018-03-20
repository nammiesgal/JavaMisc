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
import canrad.geometry.components.Feeder;
import canrad.geometry.components.JunctionDevice;
import canrad.geometry.components.Port;
import canrad.geometry.components.Segment;
import canrad.geometry.components.SiteExportViewModel;
import canrad.geometry.components.Transceiver;
import canrad.layout.models.CanradModelVisualState;
import canrad.layout.models.CellConnectionDetails;
import canrad.layout.models.FrequencyRange;
import canrad.layout.models.PortConnectionDetails;
import canrad.misc.ModelBaseTechColour;
import static canrad.misc.ModelBaseTechVisualState.techMap;
import canrad.reference.components.AntennaTypePortSegment;
import canrad.reference.components.CanradLibrary;
import canrad.reference.components.Frequency;
import canrad.reference.components.MobilesCell;
import canrad.reference.components.MobilesFrequencyBand;
import canrad.reference.components.TransceiverFuncType;
import dialog.geometry.viewmodel.ModelVisualState;
import dialog.geometry.visuals.InteractiveVisualLocation;
import dialog.geometry.visuals.ViewportStupid;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

/**
 * The TransceiverConnectivityHandler class provides functions used to trace a
 * mobiles cell from the Transceiver to the Antenna at the other end. A mobiles
 * cell is configured on the segment of a transceiver port. The trace
 * functionality follows the feeder connections up through any number of
 * Junction Devices and terminates on a segment on an antenna port. Along the
 * way, the frequency ranges on the devices are analysed to determine whether
 * the signal will continue through or will terminate.
 *
 * If the path trace lands on a Transceiver, the trace can't continue through as
 * it isn't possible to determine where or if the signal continues from there.
 *
 * @author 
 */
public class TransceiverConnectivityHandler extends CellConnectivityHelper
{
    protected ViewportStupid viewport = null;
    private List<Feeder> feederList;
    private final HashMap<Feeder, FeederTechnology> feederTechnologyMap;
    private List<Port> handledPorts;

    private final List<InteractiveVisualLocation> highlightList;
    private boolean isCellTrace;
    private List<List<PortConnectionDetails>> mappedPortsList;
    private final Set<Feeder> matchedFeederSet = new HashSet<>();
    private final SiteExportViewModel siteExportViewModel;

    /**
     *
     * @param viewModel
     */
    public TransceiverConnectivityHandler(SiteExportViewModel viewModel)
    {
        highlightList = new ArrayList<>();
        mappedPortsList = new ArrayList<>();
        feederList = new ArrayList<>();
        siteExportViewModel = viewModel;
        feederTechnologyMap = new HashMap<>();
        isCellTrace = false;
    }

    /**
     * Clear feeder trace highlights
     *
     * @param usage
     */
    public void clearHighlights(String usage)
    {
        for (InteractiveVisualLocation vl : highlightList)
        {
            vl.extractVisualStateMap(CanradModelVisualState.Standard, viewport, usage);
            vl.setVisualState(CanradModelVisualState.Standard, false, viewport);
        }
        highlightList.clear();
        feederTechnologyMap.clear();
    }

    //TODO
    public List<InteractiveVisualLocation> getHighlightList()
    {
        return highlightList;
    }

    public void setIsCellTrace(boolean value)
    {
        isCellTrace = value;
    }

    public void setSegmentVisualStates(String usage)
    {
        boolean isTechnologyView = siteExportViewModel.getCurrentSpecialVisualState().name.equals("Technology");
        boolean showRET = !isCellTrace;
        for (FeederTechnology ft : feederTechnologyMap.values())
        {
            if (showRET)
            {
                boolean hasRET = ft.feeder.getAttributes().contains("RET");
                if (hasRET)
                {
                    ModelVisualState vs = techMap.getValueForKey("RET");
                    if (vs != null)
                    {
                        InteractiveVisualLocation vl = (InteractiveVisualLocation) ft.feeder.getVisualLocation();
                        if (vl != null)
                        {
                            vl.addVisualStateMap(CanradModelVisualState.Standard, vs, viewport, usage);
                            vl.setVisualState(CanradModelVisualState.Standard, false, viewport);
                            highlightList.add(vl);
                        }
                        continue;
                    }
                }
            }
            List<MobilesCell> cellList = ft.getCellsOnFeeder(true, false);
            if (cellList.size() > 0)
            {
                ModelVisualState vs;
                if (isCellTrace && isTechnologyView)
                    vs = CanradModelVisualState.ConnectedMatched;
                else
                {
                    String highestTech = ModelBaseTechColour.getHighestTechFunction().apply(cellList);
                    vs = (highestTech != null) ? techMap.getValueForKey(highestTech) : null;
                }
                if (vs != null)
                {
                    InteractiveVisualLocation vl = (InteractiveVisualLocation) ft.feeder.getVisualLocation();
                    if (vl != null)
                    {
                        vl.addVisualStateMap(CanradModelVisualState.Standard, vs, viewport, usage);
                        vl.setVisualState(CanradModelVisualState.Standard, false, viewport);
                        highlightList.add(vl);
                    }
                }
            }
            else
            {
                cellList = ft.getCellsOnFeeder(false, true);
                if (cellList.size() > 0)
                {
                    ModelVisualState vs;
                    if (this.isCellTrace)
                        vs = CanradModelVisualState.ConnectedMismatched;
                    else
                    {
                        String highestTech = ModelBaseTechColour.getHighestTechFunction().apply(cellList);
                        vs = (highestTech != null) ? techMap.getValueForKey(highestTech) : null;
                    }
                    if (vs != null)
                    {
                        InteractiveVisualLocation vl = (InteractiveVisualLocation) ft.feeder.getVisualLocation();
                        if (vl != null)
                        {
                            vl.addVisualStateMap(CanradModelVisualState.Standard, vs, viewport, usage);
                            vl.setVisualState(CanradModelVisualState.Standard, false, viewport);
                            highlightList.add(vl);
                        }
                    }
                }
            }
        }
    }

    /**
     * Retrieves the set of mobiles cells that exist on all the segments of all
     * the ports of the given transceiver. All duplicates are filtered out and
     * the set is sorted alphabetically.
     *
     * @param transceiver The transceiver object to be analysed
     * @return the sorted set of unique mobiles cells on the transceiver
     */
    public Set<MobilesCell> retrieveTransceiverMobileCells(Transceiver transceiver)
    {
        Set<MobilesCell> cells = new TreeSet<>();

        transceiver.getPorts().forEach(port
                -> 
                {
                    Set<MobilesCell> portCells = retrieveTransceiverMobileCells(port);
                    cells.addAll(portCells);
        });
        return cells;
    }

    /**
     * Finds the segments with the given cell and traces the connectivity
     * through the connections to the antenna or until it can't continue further
     * due to either a mis-match in frequency or because the connection
     * terminates on another Transceiver.
     *
     * @param transceiver the transceiver containing the segment with the cell
     * to be traced
     * @param cell the cell to be traced
     * @param allowColourChange
     * @param usage
     * @return the list of connection details
     */
    public List<CellConnectionDetails> traceMobilesCellConnectivity(Transceiver transceiver, MobilesCell cell, boolean allowColourChange, String usage)
    {
        viewport = siteExportViewModel.getViewport(SiteExportViewModel.CONNECTIVITY_WORKSPACE_NAME);
        if (allowColourChange && highlightList.size() > 0)
            clearHighlights(usage);

        matchedFeederSet.clear();
        mappedPortsList.clear();
        feederList.clear();

        List<CellConnectionDetails> cellConnections = new ArrayList<>();

        transceiver.getSortedPorts().stream().forEach(port
                -> 
                {
                    port.getSegmentsList().stream().forEach(segment
                            -> 
                            {
                                if (segment.getMobilesCell() == cell)
                                {
                                    cellConnections.addAll(findConnectivityForSegment(port, segment));
                                    mappedPortsList = new ArrayList<>();
                                    feederList = new ArrayList<>();
                                }
                    });
        });
        return cellConnections;
    }

    private FeederTechnology addFeederTechnology(Feeder feeder)
    {
        assert feeder != null : "addFeederTechnology feeder is not null";
        FeederTechnology existing = feederTechnologyMap.get(feeder);
        if (existing == null)
            feederTechnologyMap.put(feeder, existing = new FeederTechnology(feeder));
        return existing;
    }

    private FrequencyRange extractFrequencyRangeForSegment(Segment segment, DeviceBase device)
    {
        final CanradLibrary library = CanradLibrary.getInstance();

        if (segment.getAntennaTypePortSegment() != null)
        {
            // Is an antenna segment, the frequency is on the ...
            final AntennaTypePortSegment typePortSegment = segment.getAntennaTypePortSegment();

            // Save the frequency range defined in the antenna port segment for use during the search
            return new FrequencyRange(typePortSegment.getMinFreq(), typePortSegment.getMaxFreq());
        }
        else if (segment.getTransceiverTypePortSegment() != null)
        {

            // If it has a mobiles frequency band, the frequency range is either the mobiles band (if provided) or the
            // transceiver range
            Double minFrequency, maxFrequency;
            final Integer mobilesFreqBandId = segment.getMobilesFreqBandId();
            if (mobilesFreqBandId != null)
            {
                final MobilesFrequencyBand frequencyBand = library.getMobilesFrequencyBandById(mobilesFreqBandId);
                minFrequency = frequencyBand.getMinFrequency();
                maxFrequency = frequencyBand.getMaxFrequency();
            }
            else
            {
                final Frequency trxFrequency = library.getFrequency(((Transceiver) device).getFrequencyTypeId());
                minFrequency = trxFrequency.getFreqMHz();
                maxFrequency = trxFrequency.getPartnerMHz();
            }
            return new FrequencyRange(minFrequency, maxFrequency);
        }
        return null;
    }

    private List<PortConnectionDetails> findAntennaConnectionPath(PortConnectionDetails connection)
    {
        List<PortConnectionDetails> antennaConnections = new ArrayList<>();
        handledPorts = new ArrayList<>();

        findConnectionPath(connection, antennaConnections);
        return antennaConnections;
    }

    /**
     * Get all transceiver ports for the given feeder and recursively follow any
     * junction devices.
     *
     * @param feeder
     * @param segments
     */
    private void findConnectionPath(PortConnectionDetails connection, List<PortConnectionDetails> antennaConnections)
    {
        // Keep track of the ports we have followed
        Port port = connection.getConnectedPort();
        Feeder feeder = connection.getFeeder();
        if (handledPorts.contains(port))
        {
            return;
        }
        handledPorts.add(port);
        
        if (connection.getFeeder() == null)
            return;

        if (connection.getConnectedDevice() == null)
        {
            addFeederTechnology(feeder).addSegmentMatched(connection.getSegment(), false);
            return;
        }

        // If the connected device is an Antenna, add it to the list and return
        if (connection.getConnectedDevice() instanceof Antenna)
        {
            antennaConnections.add(connection);
            return;
        }

        // If the connected device is a Transceiver, we can't trace beyond so mark as a mis-match
        if (connection.getConnectedDevice() instanceof Transceiver)
        {
            addFeederTechnology(feeder).addSegmentMatched(connection.getSegment(), false);
            return;
        }

        // If it is a junction device, find the mapped ports and call this recursively checking for a fit in frequency
        // of the device path traversed and the allowed frequency of the current junction device
        if (connection.getConnectedDevice() instanceof JunctionDevice)
        {
            JunctionDevice jd = (JunctionDevice) connection.getConnectedDevice();
            
            final List<FrequencyRange> frequencyRanges = new ArrayList<>();
            frequencyRanges.addAll(connection.getIncomingFrequencyRanges());
            List<PortConnectionDetails> mappedPorts = findMappedPorts(jd, port, frequencyRanges);
            if (!mappedPorts.isEmpty())
            {
                //System.out.println("Adding match to feeder " + connection.getFeeder().getNumber());
                connection.setPass();
                matchedFeederSet.add(connection.getFeeder());
                //These lists are populated to provide info in the cell connectivity tree view

                feederList.add(feeder);
                mappedPortsList.add(mappedPorts);
                addFeederTechnology(feeder).addSegmentMatched(connection.getSegment(), true);

                mappedPorts.stream().forEach((nextConnection)
                        -> 
                        {
                            if (nextConnection.getFeeder() != null)
                            {
                                nextConnection.setSegment(connection.getSegment());
                                findConnectionPath(nextConnection, antennaConnections);
                            }
                });
            }
            else
            {
                //System.out.println("Adding mis-match to feeder " + connection.getFeeder().getNumber());
                connection.setStop();
                addFeederTechnology(feeder).addSegmentMatched(connection.getSegment(), false);
            }
        }
    }

    /**
     * Finds the connectivity details for a single transceiver segment.
     * <p>
     * @param port
     * @param segment The details about the segment being searched
     * @param mobilesCell
     * @param connectedFunction
     * @param visualState
     * @return The list of the cell connection details for the given segment
     */
    private List<CellConnectionDetails> findConnectivityForSegment(
            Port port,
            Segment segment)
    {
        PortConnectionDetails connection = port.getPortConnectionDetails();
        if (connection == null)
        {
            return new ArrayList<>();
        }

        connection.setSegment(segment);

        /*
         * ModelVisualState modelVisualStateConnected; ModelVisualState
         * modelVisualStateMismatch;
         *
         * if (visualState == null) { modelVisualStateConnected =
         * CanradModelVisualState.ConnectedMatched; modelVisualStateMismatch =
         * CanradModelVisualState.ConnectedMismatched; } else {
         * modelVisualStateConnected = visualState; modelVisualStateMismatch =
         * visualState; }
         */
        // Save the frequency range defined in the starting port segment for use during the search
        FrequencyRange frequency = extractFrequencyRangeForSegment(segment, connection.getDevice());

        connection.getIncomingFrequencyRanges().clear();
        connection.getIncomingFrequencyRanges().add(frequency);

        // Traverse recursively through the connected devices to find the antenna ports at the other end
        List<PortConnectionDetails> antennaConnections = findAntennaConnectionPath(connection);

        // If no transceivers were found, return the empty list
        if (antennaConnections.isEmpty())
        {
            return new ArrayList<>();
        }

        MobilesCell mobilesCell = segment.getMobilesCell();
        TransceiverFuncType connectedFunction = segment.getSegmentFuncType();

        // For each port / segment returned, ensure that there is a fit between the frequency passed through the
        // Junction Devices and the segment on the transceiver.  If there is a fit, save the cell details and connected
        // function details into the cell connection details list
        List<CellConnectionDetails> cellsList = new ArrayList<>();
        for (PortConnectionDetails deviceConn : antennaConnections)
        {
            Antenna antenna = ((Antenna) deviceConn.getConnectedDevice());
            Port antennaPort = deviceConn.getConnectedPort();
            Feeder feeder = deviceConn.getFeeder();

            List<Segment> segmentList = antennaPort.getSegmentsList();
            if (segmentList.isEmpty())
            {
                addFeederTechnology(feeder).addSegmentMatched(deviceConn.getSegment(), false);
                continue;
            }
            for (Segment antennaSegment : antennaPort.getSegmentsList())
            {
                FrequencyRange antennaFrequency = extractFrequencyRangeForSegment(antennaSegment, antenna);
                // If the frequency is in the correct range, include the details for the segment

                InteractiveVisualLocation vl = (InteractiveVisualLocation) feeder.getVisualLocation();
                if (isFrequencyFit(deviceConn.getIncomingFrequencyRanges(), antennaFrequency.getMinFrequency(), antennaFrequency.getMaxFrequency()))
                {
                    // Set up the cell connection details with the connected function type
                    CellConnectionDetails details = new CellConnectionDetails();
                    details.setConnectedPort(antennaSegment.getParentPort());
                    details.setConnectedSegment(antennaSegment);
                    details.setTransceiverFuncType(connectedFunction);
                    details.setMobilesCell(mobilesCell);
                    details.setMappedPortsList(mappedPortsList);
                    matchedFeederSet.add(feeder);
                    feederList.add(feeder);
                    addFeederTechnology(feeder).addSegmentMatched(deviceConn.getSegment(), true);
                    details.setMatchedFeederList(feederList);
                    cellsList.add(details);

                }
                else if (!matchedFeederSet.contains(deviceConn.getFeeder()))
                    addFeederTechnology(feeder).addSegmentMatched(deviceConn.getSegment(), false);
            }
        }
        return cellsList;
    }

    /**
     * Retrieves the set of mobiles cells that exist on each of the segments of
     * the given transceiver port
     *
     * @param port The transceiver port object to be analysed
     * @return the set of unique mobiles cells on the port
     */
    private Set<MobilesCell> retrieveTransceiverMobileCells(Port port)
    {
        Set<MobilesCell> cells = new TreeSet<>();

        port.getSegmentsList().forEach(segment
                -> 
                {
                    final Integer cellId = segment.getMobilesCellId();
                    if (cellId != null)
                    {
                        final MobilesCell cell = CanradLibrary.getInstance().getMobilesCell(cellId);
                        if (cell != null)
                        {
                            cells.add(cell);
                        }
                    }
        });
        return cells;
    }

    private static class FeederTechnology
    {
        public final Feeder feeder;
        private final List<SegmentMatched> segments;

        public FeederTechnology(Feeder feeder)
        {
            this.feeder = feeder;
            segments = new ArrayList<>();
        }

        public void addSegmentMatched(Segment segment, boolean matched)
        {
            segments.add(new SegmentMatched(segment, matched));
        }

        public List<MobilesCell> getCellsOnFeeder(boolean includeMatched, boolean includeMismatched)
        {
            List<MobilesCell> list = new ArrayList<>();
            for (SegmentMatched segmentMatched : segments)
            {
                if (segmentMatched.matched)
                {
                    if (!includeMatched)
                        continue;
                }
                else if (!includeMismatched)
                    continue;
                MobilesCell cell = segmentMatched.segment.getMobilesCell();
                if (!list.contains(cell))
                    list.add(cell);
            }
            return list;
        }

        public static class SegmentMatched
        {
            public final Segment segment;
            public final boolean matched;

            public SegmentMatched(Segment segment, boolean matched)
            {
                this.segment = segment;
                this.matched = matched;
            }
        }

    }

}
