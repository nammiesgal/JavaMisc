/*
 * Copyright 2016 Dialog Information Technology.
 *
 * This source code is the property of Dialog Information Technology.
 * You may not use this code without the express permission of
 * the owner.
 */
package canrad.misc;

import static canrad.event.handlers.CanradBaseActionHandler.FORM_BASE_LOCATION;
import canrad.form.controllers.VisualStateLegendController;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.paint.Color;
import canrad.layout.models.CanradModelVisualState;

/**
 *
 * @author SeetoB
 */
public class VisualStateLegend
{
    private Parent root;
    private VisualStateLegendController visualStateController;
    private final ObservableList<VisualStateLegend> data = FXCollections.observableArrayList();
    
    public VisualStateLegend()
    {
        try
        {
            FXMLLoader loader = new FXMLLoader(getClass().getResource(FORM_BASE_LOCATION + "VisualStateLegend.fxml"));
            root = loader.load();
            visualStateController = (VisualStateLegendController) loader.getController();
            
        } catch (IOException ex)
        {
            Logger.getLogger(VisualStateLegend.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    public VisualStateLegendController getController() {
        return visualStateController;
    }
    public Parent getParent() {
        return root;
    }
    public void buildStatusVisualStateLegend() {
        //retrieve data from viewport
       List<Color> statusColourList = CanradModelVisualState.getStatusColours();
       List<CanradModelVisualState> statusVisualStateList = CanradModelVisualState.statusVisualStates;
     
        //access controller for pop up window
        visualStateController.setData("Status Legend", statusVisualStateList, statusColourList);
    }
    public void buildTechVisualStateLegend() {
        //retrieve data from viewport
       List<Color> statusTechList = CanradModelVisualState.getTechnologyColours();
       List<CanradModelVisualState> techVisualStateList = CanradModelVisualState.technologyVisualStates;
     
        //access controller for pop up window
        visualStateController.setData("Technology Legend", techVisualStateList,statusTechList);
    }
    
}
