import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";

import { IItem, IItemAddResult } from '@pnp/sp/items/types';
import { IAttachments } from "@pnp/sp/attachments";

// import * as strings from 'ExtensionCamlCommandSetStrings';

export interface IExtensionCamlCommandSetProperties {
  // This is an example; replace with your own properties
  //sampleTextOne: string;
}

const LOG_SOURCE: string = 'ExtensionCamlCommandSet';

export default class ExtensionCamlCommandSet extends BaseListViewCommandSet<IExtensionCamlCommandSetProperties> {

  private row : RowAccessor; // contient la ligne de la liste actuellement sélectionnée

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ExtensionCamlCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {

      if(event.selectedRows.length === 1) { // si exactement une ligne est sélectionnée alors la commande est visible
        compareOneCommand.visible = true;
        this.row = event.selectedRows[0];

      } else {
        compareOneCommand.visible = false; // sinon la commande est cachée
      }
    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> { // fonction appelée lors d'un clic sur le bouton correspondant à la commande
    switch (event.itemId) {
      case 'COMMAND_1':
                     
        const sp = spfi().using(SPFx(this.context)); // "Behavior" ou "Comportement" traduit en français, correspond au "contexte SPFx actuel"

        try { 

          // ---------------------------------------------------------------------
          // ||                         RETRIEVE DATA                           ||
          // ---------------------------------------------------------------------
          
          /*// On récupère les items qui ont la valeur "Terminée" dans leur champ "Status" 
          const items = await sp.web.lists.getByTitle("Suivi des problèmes").getItemsByCAMLQuery({
            ViewXml: '<View><Query><Where><Eq><FieldRef Name=\'Status\' /><Value Type=\'Choice\'>Terminée</Value></Eq></Where></Query></View>'
          });
          
          try { // On affiche le statut du premier item récupéré dans une boîte de dialogue, "Terminée" si tout marche correctement
            if (items.length > 0) {
              Dialog.alert(items[0].Status); // items[X].NomAttribut
            }
          } catch (e) {
            console.log("Erreur lors du traitement du résultat de la requête");
          }*/

          // ---------------------------------------------------------------------
          // ||                         CREATE DATA                             ||
          // ---------------------------------------------------------------------

          /*// Création l'objet que l'on va insérer dans la liste 
          const regex = /([0-9]{2})\/([0-9]{2})/g;
          const d = this.row.getValueByName("DateReported").replace(regex, '$2/$1');
          const date = new Date(d).toJSON();

          const properties = {
            Title : this.row.getValueByName("Title"),
            Description : this.row.getValueByName("Description"),
            Priority : this.row.getValueByName("Priority"),
            Status : this.row.getValueByName("Status"),
            AssignedtoId : this.row.getValueByName("Assignedto")[0]["id"],
            DateReported : date,
            IssueSource : {
              Url : this.row.getValueByName("IssueSource"), 
            },
            Images : JSON.stringify({
              serverUrl : this.row.getValueByName("Images")["serverUrl"],
              type : this.row.getValueByName("Images")["type"],
              fileName : this.row.getValueByName("Images")["fileName"],
              serverRelativeUrl : this.row.getValueByName("Images")["serverRelativeUrl"]
            }),
            //Attachments : this.row.getValueByName("Attachments"), // impossible de récupérer le contenu du fichier avec la ligne, on peut simplement récupérer le nombre de pièces jointes
                                                                    // il faut utiliser une autre méthode pour ajouter une pièce jointe
            IssueloggedbyId : this.row.getValueByName("Issueloggedby")[0]["id"],
            TESTCOL1 : this.row.getValueByName("TESTCOL1"),
            TEST_x0020_COL_x0020_SITE_x0020_1 : this.row.getValueByName("TEST_x0020_COL_x0020_SITE_x0020_1"),
            TEST_x0020_COL_x0020_SITE_x0020_2 : this.row.getValueByName("TEST_x0020_COL_x0020_SITE_x0020_2")
          };

          // Insertion de l'objet dans la liste
          const iar : IItemAddResult = await sp.web.lists.getByTitle("Suivi des problèmes").items.add(properties);

          // ---------------------------------------------------------------------
          // ||         BONUS : ajout d'une pièce jointe à un item              ||
          // ---------------------------------------------------------------------

          const list = sp.web.lists.getByTitle("Suivi des problèmes");
          const items = await list.getItemsByCAMLQuery({ ViewXml :
         '<View>' +
            '<Query>'+ 
              '<Where>' +
                  '<And>' +
                      '<Eq>' +
                        '<FieldRef Name=\'Title\' />' +
                        '<Value Type=\'Text\'>' + properties["Title"] + '</Value>' +                                            //  Titre / colonne problème
                      '</Eq>' +
                      '<And>' +
                        '<Eq>' +
                            '<FieldRef Name=\'Assignedto\' />' +
                            '<Value Type=\'User\'>' + this.row.getValueByName("Assignedto")[0]["title"] + '</Value>' +          // nom d'affichage de l'utilisateur
                        '</Eq>' +
                        '<Eq>' +
                            '<FieldRef Name=\'DateReported\' />' +
                            '<Value Type=\'DateTime\'>' + '2022-05-03' + '</Value>' +   // string d'une date sérialisée au format JSON
                        '</Eq>' +
                      '</And>' +
                  '</And>' +
                '</Where>' +
            '</Query>' +
          '</View>'});

          console.log(items); //debug
          
          for(let i=0; i<items.length; i++) {
            (async () => {
              const item : IItem = list.items.getById(items[i]["Id"]);
              await item.attachmentFiles.add("FichierTest.txt", "Contenu du fichier"); // Contenu = string, Blob ou ArrayBuffer
              console.log("Pièce jointe avec succès");
            })().catch(console.log);
          }*/
          
          // ---------------------------------------------------------------------
          // ||                         UPDATE DATA                             ||
          // ---------------------------------------------------------------------
          
          /*const items: any[] = await sp.web.lists.getByTitle("Suivi des problèmes").items.filter("Title eq 'Pb Test 3'")();

          if (items.length > 0) {
            const updatedItem = await sp.web.lists.getByTitle("Suivi des problèmes").items.getById(items[0].Id).update({
              Title: "Updated Title",
            });
          }*/

          // ---------------------------------------------------------------------
          // ||                         DELETE DATA                             ||
          // ---------------------------------------------------------------------

          const items: any[] = await sp.web.lists.getByTitle("Suivi des problèmes").items.filter("Title eq 'test'")();

          if (items.length > 0) {
            const updatedItem = await sp.web.lists.getByTitle("Suivi des problèmes").items.getById(items[0].Id).delete();
          }

        } catch (e) {
          console.error("Erreur lors de l'accès à l'API");
        }
        break;

      default:
        throw new Error('Unknown command');
    }
  }
}
