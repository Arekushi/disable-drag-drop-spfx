import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "DragDropDeactivatorApplicationCustomizerStrings";
import { IDragDropDeactivatorApplicationCustomizerProperties } from "./DragDropDeactivatorApplicationCustomizerProperties";

import { ListService } from "../../services/ListService";
import { LimitedResourcesListColumns as columns } from "../../enums/LimitedResourcesListColumnsEnum";

const LOG_SOURCE: string = "DragDropDeactivatorApplicationCustomizer";


/** A Custom Action which can be run during execution of a Client Side Application */
export default class DragDropDeactivatorApplicationCustomizer
    extends BaseApplicationCustomizer<IDragDropDeactivatorApplicationCustomizerProperties> {

    private limitedResourceService: ListService;

    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

        this.limitedResourceService = new ListService(
            this.properties.listName,
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl
        );

        const pageList = this.context.pageContext.list;

        // Check if the current page is a list/library
        if (pageList) {
            this.limitedResourceService
                .getListItems(
                    `$filter=`
                        + `${columns.RECURSO_DESABILITADO} eq '${this.properties.resourceName}' `
                        + `and ${columns.TITLE} eq '${pageList.title}'`
                )
                .then((data) => {
                    console.log(data);

                    if (data.value.length > 0) {
                        this.listenCancelDragDropEvent();
                    }
                })
                .catch((err) => {
                    Log.error(strings.MessageOnDragDrop, err);
                });
        }

        return Promise.resolve();
    }

    private listenCancelDragDropEvent(): void {
        document.addEventListener("drop", this.cancelDropEvent, false);
        document.addEventListener("dragover", this.cancelDragEvent, false);
        document.addEventListener("dragenter", this.cancelDragEvent, false);
    }

    private cancelDropEvent(event: Event): void {
        event.preventDefault();
        event.stopPropagation();

        Dialog.alert(strings.MessageOnDragDrop)
            .catch((err) => {
                Log.error(strings.MessageOnDragDrop, err);
            })
    }

    private cancelDragEvent(event: Event): void {
        event.preventDefault();
        event.stopPropagation();
    }

}
