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

import canrad.celltrace.TransceiverConnectivityHandler;
import canrad.geometry.components.Feeder;
import canrad.geometry.components.Port;
import canrad.geometry.components.SiteExportViewModel;
import canrad.geometry.components.Transceiver;
import canrad.layout.models.CanradModelVisualState;
import canrad.reference.components.CanradLibrary;
import canrad.reference.components.MobilesCell;
import dialog.geometry.components.Component;
import dialog.geometry.viewmodel.ModelVisualState;
import dialog.geometry.visuals.InteractiveVisualLocation;
import dialog.geometry.visuals.ViewportStupid;
import dialog.utilities.KeyValueMap;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

/**
 *
 * @author SeetoB
 */
public class ModelBaseTechVisualState
{
    public final static KeyValueMap<String, ModelVisualState> techMap;
    private final TransceiverConnectivityHandler connectivityHandler;
   
    static
    {
        techMap = new KeyValueMap<>();
        addTechnology("RET");
        CanradLibrary.getInstance().getMobilesTechnologyTypes().forEach(techType
                -> 
                {
                    String techTypeCode = techType.getTypeCode();
                    addTechnology(techTypeCode);
        });
    }
    
    private static void addTechnology(String code)
    {
        ModelVisualState visualState = ModelVisualState.getModelVisualStateByName(code);
        if (visualState != null)
        {
            techMap.addKeyValuePair(code, visualState);
        }
    }
    
    public static ModelVisualState getVisualState(String techCode)
    {
        return techMap.getValueForKey(techCode);        
    }

    public ModelBaseTechVisualState(SiteExportViewModel siteViewModel)
    {
        connectivityHandler = new TransceiverConnectivityHandler(siteViewModel);
    }
    
    public TransceiverConnectivityHandler getConnectivityHandler()
    {
        return connectivityHandler;
    }

    public Function<Component, ModelVisualState> getTechVisualStateFunction()
    {
        return new Function<Component, ModelVisualState>()
        {
            @Override
            public ModelVisualState apply(Component component)
            {
                if (component instanceof Port)
                {
                    if (((Port) component).getParentDevice() instanceof Transceiver)
                    {
                        Transceiver parentTransceiver = (Transceiver) ((Port) component).getParentDevice();
                        List<MobilesCell> mobilesCells = new ArrayList<MobilesCell>(connectivityHandler.retrieveTransceiverMobileCells(parentTransceiver));

                        if (mobilesCells.size() > 0)
                        {   
                            mobilesCells.forEach(cell
                                    -> 
                                    {
                                        if (cell != null)
                                        {
                                            connectivityHandler.traceMobilesCellConnectivity(parentTransceiver, cell, true, "Technology");
                                        }
                                        connectivityHandler.getHighlightList().clear();
                            });
                        }
                    }
                    
                }
                else if (component instanceof Feeder)
                {
                    Feeder feeder = (Feeder)component;
                    if (feeder.getAttributes().contains("RET"))
                    {
                        InteractiveVisualLocation vl = (InteractiveVisualLocation) feeder.getVisualLocation();
                        ViewportStupid viewport = vl.getContainerVisualLocations().getViewport(SiteExportViewModel.CONNECTIVITY_WORKSPACE_NAME);
                        ModelVisualState vsRET = ModelBaseTechVisualState.getVisualState("RET");
                        vl.addVisualStateMap(CanradModelVisualState.Standard, vsRET, viewport, "Technology");
                        vl.setVisualState(CanradModelVisualState.Standard, false, viewport);
                    }
                }
                
                return ModelVisualState.Standard;
            }
        };
    }
}
