import { IThemeRules, ThemeGenerator, themeRulesStandardCreator, } from 'office-ui-fabric-react/lib/ThemeGenerator';
import { IColor } from 'office-ui-fabric-react/lib/utilities/color/interfaces';
import { getColorFromString } from 'office-ui-fabric-react/lib/utilities/color/getColorFromString';

export default class ThemeProvider {

    private themeRules: IThemeRules;

    constructor() {
        const themeRules = themeRulesStandardCreator();

        ThemeGenerator.insureSlots(this.themeRules, false);

        ThemeGenerator.setSlot(themeRules.backgroundColor, getColorFromString('#ffffff') || '', false, true, true);
        ThemeGenerator.setSlot(themeRules.foregroundColor, getColorFromString('#000000') || '', false, true, true);

        this.themeRules = themeRules;
    }

    public getThemeForColor(hexColor: string): void {

        let col: IColor = { str: '#0078D4', r: 0, g: 120, b: 212, hex: '0078d4', h: 0, s: 0, v: 0 };

        const newColor: IColor = getColorFromString(hexColor) || col;

        const themeRules = this.themeRules;

        ThemeGenerator.setSlot(themeRules.primaryColor, newColor.str || '', false, true, true);

        this.themeRules = themeRules;

        return ThemeGenerator.getThemeAsJson(this.themeRules);
    }
}