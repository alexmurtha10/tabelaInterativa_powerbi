/*
 * settings.ts
 *
 * Define as opcoes exibidas no painel de formatacao do Power BI
 * para o visual de matriz por periodo.
 *
 * Todas as cores do visual sao configuraveis pelo usuario:
 * - Titulo (texto e fundo);
 * - Tabela (cabecalhos, linhas, texto e bordas);
 * - Totais (linha Total Mensal e coluna Total Geral);
 * - Popup de detalhes (fundo, texto, cabecalho, bordas e overlay).
 */

"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Configuracoes do titulo do visual.
 */
class TitleCardSettings extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Exibir titulo",
        value: true
    });

    text = new formattingSettings.TextInput({
        name: "text",
        displayName: "Texto do titulo",
        value: "Matriz por Periodo",
        placeholder: "Digite o titulo do visual"
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Tamanho do texto",
        value: 13
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Cor do texto",
        value: { value: "#FFFFFF" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Cor do fundo",
        value: { value: "#2E4153" }
    });

    name: string = "title";
    displayName: string = "Titulo";

    slices: Array<FormattingSettingsSlice> = [
        this.show,
        this.text,
        this.fontSize,
        this.fontColor,
        this.backgroundColor
    ];
}

/**
 * Configuracoes visuais da tabela.
 */
class TableCardSettings extends FormattingSettingsCard {
    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Tamanho do texto",
        value: 12
    });

    headerFontColor = new formattingSettings.ColorPicker({
        name: "headerFontColor",
        displayName: "Cor do texto do cabecalho",
        value: { value: "#FFFFFF" }
    });

    headerBackgroundColor = new formattingSettings.ColorPicker({
        name: "headerBackgroundColor",
        displayName: "Cor do cabecalho",
        value: { value: "#3F5B70" }
    });

    groupHeaderBackgroundColor = new formattingSettings.ColorPicker({
        name: "groupHeaderBackgroundColor",
        displayName: "Cor do cabecalho dos grupos",
        value: { value: "#374E65" }
    });

    oddRowColor = new formattingSettings.ColorPicker({
        name: "oddRowColor",
        displayName: "Cor de fundo das linhas impares",
        value: { value: "#FFFFFF" }
    });

    oddRowFontColor = new formattingSettings.ColorPicker({
        name: "oddRowFontColor",
        displayName: "Cor do texto das linhas impares",
        value: { value: "#333333" }
    });

    evenRowColor = new formattingSettings.ColorPicker({
        name: "evenRowColor",
        displayName: "Cor de fundo das linhas pares",
        value: { value: "#F2F2F2" }
    });

    evenRowFontColor = new formattingSettings.ColorPicker({
        name: "evenRowFontColor",
        displayName: "Cor do texto das linhas pares",
        value: { value: "#333333" }
    });

    gridColor = new formattingSettings.ColorPicker({
        name: "gridColor",
        displayName: "Cor das bordas",
        value: { value: "#CCCCCC" }
    });

    name: string = "table";
    displayName: string = "Tabela";

    slices: Array<FormattingSettingsSlice> = [
        this.fontSize,
        this.headerFontColor,
        this.headerBackgroundColor,
        this.groupHeaderBackgroundColor,
        this.oddRowColor,
        this.oddRowFontColor,
        this.evenRowColor,
        this.evenRowFontColor,
        this.gridColor
    ];
}

/**
 * Configuracoes da linha Total Mensal e da coluna Total Geral.
 */
class TotalsCardSettings extends FormattingSettingsCard {
    rowBackgroundColor = new formattingSettings.ColorPicker({
        name: "rowBackgroundColor",
        displayName: "Cor de fundo da linha Total Mensal",
        value: { value: "#374E65" }
    });

    rowFontColor = new formattingSettings.ColorPicker({
        name: "rowFontColor",
        displayName: "Cor do texto da linha Total Mensal",
        value: { value: "#FFFFFF" }
    });

    columnBackgroundColor = new formattingSettings.ColorPicker({
        name: "columnBackgroundColor",
        displayName: "Cor de fundo da coluna Total Geral",
        value: { value: "#EAEFF3" }
    });

    columnFontColor = new formattingSettings.ColorPicker({
        name: "columnFontColor",
        displayName: "Cor do texto da coluna Total Geral",
        value: { value: "#2E4153" }
    });

    cornerBackgroundColor = new formattingSettings.ColorPicker({
        name: "cornerBackgroundColor",
        displayName: "Cor de fundo do total geral (canto)",
        value: { value: "#2E4153" }
    });

    cornerFontColor = new formattingSettings.ColorPicker({
        name: "cornerFontColor",
        displayName: "Cor do texto do total geral (canto)",
        value: { value: "#FFFFFF" }
    });

    name: string = "totals";
    displayName: string = "Totais";

    slices: Array<FormattingSettingsSlice> = [
        this.rowBackgroundColor,
        this.rowFontColor,
        this.columnBackgroundColor,
        this.columnFontColor,
        this.cornerBackgroundColor,
        this.cornerFontColor
    ];
}

/**
 * Configuracoes de cores do popup de detalhes.
 */
class PopupCardSettings extends FormattingSettingsCard {
    overlayColor = new formattingSettings.ColorPicker({
        name: "overlayColor",
        displayName: "Cor do fundo escurecido",
        value: { value: "#000000" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Cor de fundo do popup",
        value: { value: "#2E4153" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Cor do texto do popup",
        value: { value: "#FFFFFF" }
    });

    headerBackgroundColor = new formattingSettings.ColorPicker({
        name: "headerBackgroundColor",
        displayName: "Cor do cabecalho do popup",
        value: { value: "#3F5B70" }
    });

    rowAltColor = new formattingSettings.ColorPicker({
        name: "rowAltColor",
        displayName: "Cor das linhas alternadas do popup",
        value: { value: "#364D60" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Cor das bordas do popup",
        value: { value: "#4E6E85" }
    });

    name: string = "popup";
    displayName: string = "Popup de Detalhes";

    slices: Array<FormattingSettingsSlice> = [
        this.overlayColor,
        this.backgroundColor,
        this.fontColor,
        this.headerBackgroundColor,
        this.rowAltColor,
        this.borderColor
    ];
}

/**
 * Modelo principal do painel de formatacao.
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    titleCard = new TitleCardSettings();
    tableCard = new TableCardSettings();
    totalsCard = new TotalsCardSettings();
    popupCard = new PopupCardSettings();

    cards = [
        this.titleCard,
        this.tableCard,
        this.totalsCard,
        this.popupCard
    ];
}
