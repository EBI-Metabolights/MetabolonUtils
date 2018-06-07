/*
 * EMBL-EBI MetaboLights - https://www.ebi.ac.uk/metabolights
 * Metabolomics Team
 *
 *  European Bioinformatics Institute (EMBL-EBI), European Molecular Biology Laboratory, Wellcome Genome Campus, Hinxton, Cambridge CB10 1SD, United Kingdom
 *
 *  Last modified: 2018-May-10
 *  Modified by:   kenneth
 *
 *  Copyright 2018 EMBL - European Bioinformatics Institute
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
 */

package uk.ac.ebi.metabolights.utils.metabolonutils;

import org.isatools.plugins.metabolights.assignments.actions.AutoCompletionAction;
import org.isatools.plugins.metabolights.assignments.model.Metabolite;
import uk.ac.ebi.chebi.webapps.chebiWS.client.ChebiWebServiceClient;
import uk.ac.ebi.chebi.webapps.chebiWS.model.*;

import javax.xml.namespace.QName;
import java.net.MalformedURLException;
import java.net.URL;

public class SearchUtils {

    //ChEBI WS stuff
    private final String chebiWSUrl = "http://www.ebi.ac.uk/webservices/chebi/2.0/webservice?wsdl";
    private ChebiWebServiceClient chebiWS;

    public Metabolite getMetaboliteInformation(String identifier, String metaboliteName){
        // search by compound name
        Metabolite met = new Metabolite();

        //       if (metabolite != null && identifier == null)
        //           met = AutoCompletionAction.getMetaboliteFromMetaboLightWS(AutoCompletionAction.DESCRIPTION_COL_NAME, metabolite);

        String chebiId = identifier, dataType = "externalId";

        if (identifier != null) {

            if (identifier.toLowerCase().contains("hmdb")) {
                String newHmdbId = identifier;

                //Change HMDB06029 to HMDB0006029
                if (identifier.length() <= 9) // Old HMDB namespace
                    newHmdbId = identifier.replace("HMDB0","HMDB000");

                try {
                    LiteEntityList chebiList = getChebiEntity(newHmdbId, dataType);

                    for (LiteEntity le: chebiList.getListElement()){
                        chebiId = le.getChebiId();
                        break; //Only want the first entity
                    }

                } catch (ChebiWebServiceFault_Exception e) {
                    e.printStackTrace();
                }

            }
            return AutoCompletionAction.getMetaboliteFromMetaboLightWS(AutoCompletionAction.IDENTIFIER_COL_NAME, chebiId);
        }

        if (metaboliteName != null) {

                try {

                    LiteEntityList chebiList = getChebiEntity(metaboliteName, "compoundName");

                    for (LiteEntity le: chebiList.getListElement()){
                        chebiId = le.getChebiId();
                        break; //Only want the first entity. Should only be one anyway.
                    }

                } catch (ChebiWebServiceFault_Exception e) {
                    e.printStackTrace();
                }

                if (chebiId != null)
                    return AutoCompletionAction.getMetaboliteFromMetaboLightWS(AutoCompletionAction.IDENTIFIER_COL_NAME, chebiId);
        }


        return met;
    }

    public ChebiWebServiceClient getChebiWS() {
        if (chebiWS == null)
            try {
                System.out.println("Starting a new instance of the ChEBI ChebiWebServiceClient");
                chebiWS = new ChebiWebServiceClient(new URL(chebiWSUrl),new QName("https://www.ebi.ac.uk/webservices/chebi",	"ChebiWebServiceService"));
            } catch (MalformedURLException e) {
                System.out.println("Error instanciating a new ChebiWebServiceClient "+ e.getMessage());
            }
        return chebiWS;
    }

    private LiteEntityList getChebiEntity(String searchTerm, String dataType) throws ChebiWebServiceFault_Exception {

        if (dataType.equals("compoundName")) {
            LiteEntityList liteEntityList = getChebiWS().getLiteEntity(searchTerm, SearchCategory.CHEBI_NAME, 1, StarsCategory.ALL);

            if (liteEntityList.getListElement().isEmpty())
                liteEntityList = getChebiWS().getLiteEntity(searchTerm, SearchCategory.ALL_NAMES, 1, StarsCategory.ALL);

            return liteEntityList;
        }

        if (dataType.equals("externalId"))
            return getChebiWS().getLiteEntity(searchTerm, SearchCategory.DATABASE_LINK_REGISTRY_NUMBER_CITATION, 1, StarsCategory.ALL);

        return null;
    }
}
