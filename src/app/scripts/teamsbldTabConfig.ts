import {TeamsTheme} from './theme';

/**
 * Implementation of Teamsbld Tab configuration page
 */
export class teamsbldTabConfigure {
    constructor() {
        microsoftTeams.initialize();

        microsoftTeams.getContext((context:microsoftTeams.Context) => {
            TeamsTheme.fix(context);
            let val = <HTMLInputElement>document.getElementById("data");
            if (context.entityId) {
                val.value = context.entityId;
            }
            this.setValidityState(true);
        });
		
        microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {

            let val = <HTMLInputElement>document.getElementById("data");
			// Calculate host dynamically to enable local debugging
			let host = "https://" + window.location.host;
			let defaultTabName: string = `Teamsbld`;
			// Upper case first letter of tab name
			defaultTabName = defaultTabName.charAt(0).toUpperCase() + defaultTabName.slice(1);
            microsoftTeams.settings.setSettings({
                contentUrl: host + "/teamsbldTabTab.html?data=",
                suggestedDisplayName: defaultTabName,
                removeUrl: host + "/teamsbldTabRemove.html",
				entityId: val.value
            });

            saveEvent.notifySuccess();
        });
    }
    public setValidityState(val: boolean) {
        microsoftTeams.settings.setValidityState(val);
    }
}