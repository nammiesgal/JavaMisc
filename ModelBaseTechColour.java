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

import canrad.layout.models.CanradModelVisualState;
import canrad.reference.components.MobilesCell;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;
import java.util.OptionalInt;
import java.util.function.Function;
import java.util.stream.IntStream;

/**
 *
 * @author SeetoB
 */
public class ModelBaseTechColour
{
    public static Function<List<MobilesCell>, String> getHighestTechFunction()
    {
        return new Function<List<MobilesCell>, String>()
        {
            @Override
            public String apply(List<MobilesCell> mobileCellList)
            {
                List<TechOrder> techOrderList = new ArrayList();
                List<CanradModelVisualState> techColourListOrder = CanradModelVisualState.technologyVisualStates;

                mobileCellList.stream().forEach(cell -> 
                        {
                            String techCode = cell.getTechnology();
                            if (techCode != null)
                            {
                                //Find technology order in list of technologies by using the index of list.
                                OptionalInt index = IntStream.range(0, techColourListOrder.size()).filter(i -> techColourListOrder.get(i).name.equalsIgnoreCase(techCode)).findFirst();
                                //Only process known technologies
                                if (index.isPresent())
                                {
                                    //Create TechOrder object to store the technology and its order.  Store in a list.
                                    techOrderList.add(new TechOrder(techCode, index.getAsInt()));
                                }

                            }
                        });
                //Find highest technology from list of TechOrder objects
                if (techOrderList.size() > 1)
                {
                    //Do comparison to get highest tech
                    Comparator<TechOrder> byTechOrder = (t1, t2) -> Integer.compare(
                            t1.getTechOrder(), t2.getTechOrder());

                    //Sort the list and retrieve first element in list as that is the higher priority tech
                    Optional<TechOrder> obj = techOrderList.stream().sorted(byTechOrder).findFirst();
                    return(obj.get().getTechCode());

                }
                else if (techOrderList.size() == 1)
                {
                    return(techOrderList.get(0).getTechCode());
                }
                else {
                    return null;
                }
            }
        };
    }
 
    public static class TechOrder
    {

        private final String techCode;
        private final Integer techOrder;

        private TechOrder(String technologyCode, Integer order)
        {
            techCode = technologyCode;
            techOrder = order;
        }

        public String getTechCode()
        {
            return techCode;
        }

        public Integer getTechOrder()
        {
            return techOrder;
        }
    }
}
