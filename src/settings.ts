/*
 * settings.ts
 *
 * Define as opcoes exibidas no painel de formatacao do Power BI
 * para o visual de matriz por periodo.
 *
 * Permite configurar:
 * - exibicao e texto do titulo;
 * - tamanho e cores do titulo;
 * - tamanho do texto da tabela;
 * - cores dos cabecalhos;
 * - cores das linhas;
 * - cor das bordas.
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
        value: {
            value: "#FFFFFF"
        }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Cor do fundo",
        value: {
            value: "#2E4153"
        }
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
        value: {
            value: "#FFFFFF"
        }
    });

    headerBackgroundColor = new formattingSettings.ColorPicker({
        name: "headerBackgroundColor",
        displayName: "Cor do cabecalho",
        value: {
            value: "#3F5B70"
        }
    });

    groupHeaderBackgroundColor = new formattingSettings.ColorPicker({
        name: "groupHeaderBackgroundColor",
        displayName: "Cor do cabecalho dos grupos",
        value: {
            value: "#374E65"
        }
    });

    oddRowColor = new formattingSettings.ColorPicker({
        name: "oddRowColor",
        displayName: "Cor das linhas impares",
        value: {
            value: "#FFFFFF"
        }
    });

    evenRowColor = new formattingSettings.ColorPicker({
        name: "evenRowColor",
        displayName: "Cor das linhas pares",
        value: {
            value: "#F2F2F2"
        }
    });

    gridColor = new formattingSettings.ColorPicker({
        name: "gridColor",
        displayName: "Cor das bordas",
        value: {
            value: "#CCCCCC"
        }
    });

    name: string = "table";
    displayName: string = "Tabela";

    slices: Array<FormattingSettingsSlice> = [
        this.fontSize,
        this.headerFontColor,
        this.headerBackgroundColor,
        this.groupHeaderBackgroundColor,
        this.oddRowColor,
        this.evenRowColor,
        this.gridColor
    ];
}

/**
 * Modelo principal do painel de formatacao.
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    titleCard = new TitleCardSettings();
    tableCard = new TableCardSettings();

    cards = [
        this.titleCard,
        this.tableCard
    ];
}
