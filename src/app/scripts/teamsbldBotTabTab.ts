import { TeamsTheme } from './theme';

/**
 * Implementation of Bot pinned tab: Teamsbld Bot Tab
 */
export class teamsbldBotTabTab {
    constructor() {
        microsoftTeams.initialize();
        TeamsTheme.fix();
    }
    public doStuff() {
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            var a = document.getElementById('app');
            if (a) {
               // do something
            }
        });
    }

    getParameterByName(name: string, url?: string): string {
        if (!url) {
            url = window.location.href;
        }
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return '';
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

}