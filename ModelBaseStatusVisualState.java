/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package canrad.misc;

import canrad.geometry.components.ModelBase;
import canrad.geometry.components.Port;
import canrad.layout.components.FeederEndLayout;
import canrad.layout.components.FeederSegment;
import canrad.reference.components.CanradLibrary;
import dialog.geometry.components.Component;
import dialog.geometry.viewmodel.ModelVisualState;
import dialog.utilities.KeyValueMap;
import java.util.function.Function;

/**
 *
 * @author SeetoB
 */
public class ModelBaseStatusVisualState
{

    public final static KeyValueMap<String, ModelVisualState> statusMap;

    static
    {
        statusMap = new KeyValueMap<>();
        for (String statusCode : CanradLibrary.getInstance().getStatuses())
        {
            String origStatusCode = statusCode;
            while (statusCode.contains(" "))
            {
                statusCode = statusCode.replace(' ', '_');
            }
            String visualStateName = "Status_" + statusCode;
            ModelVisualState visualState = ModelVisualState.getModelVisualStateByName(visualStateName);
            statusMap.addKeyValuePair(origStatusCode, visualState);
        }
    }

    public static Function<Component, ModelVisualState> getStatusVisualStateFunction()
    {
        return (Component component)
                -> 
                {
                    String statusCode;
                    if (ModelBase.class.isAssignableFrom(component.getClass()))
                    {
                        statusCode = ((ModelBase) component).getStatusCode();

                    }
                    else if (component instanceof FeederEndLayout)
                    {
                        statusCode = ((FeederEndLayout) component).getFeederLayout().getFeeder().getStatusCode();

                    }
                    else if (component instanceof FeederSegment)
                    {
                        statusCode = ((FeederSegment) component).getFeeder().getStatusCode();

                    }
                    else if (component instanceof Port) {
                        statusCode = ((Port) component).getParentDevice().getStatusCode();
                    }
                    else
                    {
                        return ModelVisualState.Standard;
                    }
                    ModelVisualState vs = statusMap.getValueForKey(statusCode);
                    if (vs != null)
                    {
                        return vs;
                    }
                    return ModelVisualState.Standard;
        };
    }
}
