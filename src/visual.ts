/*
 * visual.ts
 *
 * Renderiza uma matriz dinamica no Power BI.
 *
 * Estrutura:
 * - Colunas informativas independentes a esquerda;
 * - Colunas mensais dinamicas a direita;
 * - Coluna "Total Geral" a direita, com a soma de cada linha;
 * - Linha "Total Mensal" fixa no rodape, com a soma de cada mes;
 * - Uma medida ou campo exibido em cada celula mensal;
 * - Detalhamento opcional ao clicar em uma celula;
 * - Titulo configuravel pelo painel de formatacao;
 * - Todas as cores sao configuraveis pelo painel de formatacao;
 * - Exportacao dos dados apresentados para CSV.
 *
 * O campo configurado em "Data" pode ser uma data real ou um texto
 * no formato jan/26, fev/26, janeiro/2026, 01/2026 etc.
 */

"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import DataViewTable = powerbi.DataViewTable;
import DataViewTableRow = powerbi.DataViewTableRow;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import FormattingModel = powerbi.visuals.FormattingModel;

import { VisualFormattingSettingsModel } from "./settings";

/**
 * CSS estrutural do visual.
 *
 * Este CSS controla apenas layout, posicionamento e comportamento
 * (sticky, scroll, espacamentos). Nenhuma cor fixa e definida aqui;
 * todas as cores sao aplicadas via estilo inline a partir das
 * configuracoes escolhidas no painel de formatacao.
 */
const VISUAL_CSS = `
.fperiod-wrap {
    display: flex;
    flex-direction: column;
    height: 100%;
    overflow: hidden;
    font-family: Calibri, "Segoe UI", sans-serif;
    background: transparent;
}

.fperiod-title {
    flex-shrink: 0;
    box-sizing: border-box;
    width: 100%;
    padding: 7px 10px;
    text-align: center;
    font-weight: bold;
}

.fperiod-title-hidden {
    display: none;
}

.fperiod-scroll {
    flex: 1;
    min-height: 0;
    overflow: auto;
}

.fperiod-table {
    border-collapse: separate;
    border-spacing: 0;
    width: 100%;
    min-width: max-content;
    table-layout: auto;
}

.fperiod-th-group {
    position: sticky;
    top: 0;
    z-index: 12;
    padding: 6px 8px;
    text-align: center;
    font-weight: bold;
    white-space: nowrap;
    text-transform: uppercase;
    letter-spacing: 0.3px;
}

.fperiod-th-info,
.fperiod-th-period,
.fperiod-th-total {
    position: sticky;
    z-index: 11;
    padding: 6px 8px;
    text-align: center;
    font-weight: bold;
    white-space: nowrap;
}

.fperiod-th-info {
    min-width: 90px;
}

.fperiod-th-period,
.fperiod-th-total {
    min-width: 64px;
}

.fperiod-td-info,
.fperiod-td-period,
.fperiod-td-total {
    box-sizing: border-box;
    padding: 6px 8px;
    white-space: nowrap;
}

.fperiod-td-info {
    text-align: left;
}

.fperiod-td-period {
    min-width: 64px;
    text-align: center;
    font-weight: normal;
}

.fperiod-td-total {
    min-width: 64px;
    text-align: center;
    font-weight: bold;
}

.fperiod-td-period.clickable {
    cursor: pointer;
}

.fperiod-td-period.clickable:hover {
    filter: brightness(0.94);
}

/* Rodape fixo (Total Mensal) */
.fperiod-tfoot-cell {
    position: sticky;
    bottom: 0;
    z-index: 11;
    box-sizing: border-box;
    padding: 6px 8px;
    text-align: center;
    font-weight: bold;
    white-space: nowrap;
}

.fperiod-tfoot-label {
    text-align: left;
}

.fperiod-tfoot-corner {
    z-index: 13;
}

.fperiod-empty {
    padding: 40px;
    text-align: center;
    color: #888888;
    font-size: 13px;
}

.fperiod-overlay {
    display: none;
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    z-index: 99998;
}

.fperiod-modal {
    display: none;
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    box-sizing: border-box;
    width: 720px;
    max-width: 95vw;
    max-height: 80vh;
    overflow: auto;
    padding: 14px;
    border-radius: 10px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.5);
    z-index: 99999;
    font-family: Calibri, "Segoe UI", sans-serif;
    font-size: 12px;
}

.fperiod-modal-header {
    display: flex;
    align-items: center;
    margin-bottom: 10px;
}

.fperiod-modal-title {
    flex: 1;
    font-size: 13px;
    font-weight: bold;
}

.fperiod-modal-close {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 26px;
    height: 26px;
    padding: 0;
    border: none;
    border-radius: 50%;
    font-size: 14px;
    cursor: pointer;
}

.fperiod-modal-table {
    border-collapse: collapse;
    width: 100%;
}

.fperiod-modal-table th {
    position: sticky;
    top: 0;
    padding: 5px 8px;
    font-weight: bold;
    white-space: nowrap;
}

.fperiod-modal-table td {
    padding: 4px 7px;
    white-space: nowrap;
}
`;

/**
 * Aplica opacidade a uma cor hexadecimal, retornando um valor rgba.
 *
 * Usado para o fundo escurecido do popup a partir de uma cor solida.
 */
function hexToRgba(hex: string, alpha: number): string {
    const clean = hex.replace("#", "").trim();

    const full =
        clean.length === 3
            ? clean.split("").map(c => c + c).join("")
            : clean;

    const red = parseInt(full.substring(0, 2), 16);
    const green = parseInt(full.substring(2, 4), 16);
    const blue = parseInt(full.substring(4, 6), 16);

    if (isNaN(red) || isNaN(green) || isNaN(blue)) {
        return `rgba(0, 0, 0, ${alpha})`;
    }

    return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

/**
 * Insere ou atualiza uma tag de estilo no documento.
 */
function injectStyle(id: string, css: string): void {
    let element = document.getElementById(id) as HTMLStyleElement | null;

    if (!element) {
        element = document.createElement("style");
        element.id = id;
        document.head.appendChild(element);
    }

    element.textContent = css;
}

/**
 * Escapa valores antes de inseri-los no HTML.
 */
function escapeHtml(value: unknown): string {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/**
 * Retorna somente os numeros contidos em um texto.
 */
function onlyDigits(value: string): string {
    return value.replace(/\D/g, "");
}

/**
 * Converte um valor para numero, ou retorna null quando nao for numerico.
 */
function toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === "") {
        return null;
    }

    const parsed = typeof value === "number" ? value : Number(value);

    return isNaN(parsed) ? null : parsed;
}

/**
 * Converte um ano com dois digitos para quatro digitos.
 *
 * Exemplo:
 * 26 -> 2026
 * 99 -> 1999
 */
function normalizeYear(year: number): number {
    if (year >= 100) {
        return year;
    }

    return year <= 69 ? 2000 + year : 1900 + year;
}

/**
 * Retorna o numero do mes a partir de um nome ou abreviacao.
 */
function getMonthNumber(value: string): number | null {
    const normalized = value
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toLowerCase()
        .substring(0, 3);

    const months: Record<string, number> = {
        jan: 1,
        fev: 2,
        mar: 3,
        abr: 4,
        mai: 5,
        jun: 6,
        jul: 7,
        ago: 8,
        set: 9,
        out: 10,
        nov: 11,
        dez: 12
    };

    return months[normalized] ?? null;
}

/**
 * Representa um periodo utilizado para criar uma coluna dinamica.
 */
interface PeriodInfo {
    key: string;
    label: string;
    sortValue: number;
}

/**
 * Cria a chave mensal a partir do campo Data.
 *
 * A funcao aceita:
 * - objeto Date;
 * - data ISO;
 * - timestamp;
 * - dd/MM/yyyy;
 * - MM/yyyy;
 * - jan/26;
 * - janeiro/2026;
 *
 * Quando o valor nao puder ser interpretado como mes e ano,
 * o proprio texto sera usado como coluna.
 */
function createPeriodInfo(value: unknown): PeriodInfo | null {
    if (value === null || value === undefined || value === "") {
        return null;
    }

    const monthNames = [
        "",
        "jan",
        "fev",
        "mar",
        "abr",
        "mai",
        "jun",
        "jul",
        "ago",
        "set",
        "out",
        "nov",
        "dez"
    ];

    let month: number | null = null;
    let year: number | null = null;

    if (value instanceof Date && !isNaN(value.getTime())) {
        month = value.getMonth() + 1;
        year = value.getFullYear();
    } else if (typeof value === "number") {
        const numericDate = new Date(value);

        if (!isNaN(numericDate.getTime())) {
            month = numericDate.getMonth() + 1;
            year = numericDate.getFullYear();
        }
    } else {
        const raw = String(value).trim();

        // Formato ISO: 2026-01-01 ou 2026-01-01T00:00:00
        const isoMatch = raw.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?/);

        if (isoMatch) {
            year = parseInt(isoMatch[1], 10);
            month = parseInt(isoMatch[2], 10);
        }

        // Formato brasileiro com dia: 01/01/2026 ou 01-01-2026
        if (month === null || year === null) {
            const brDateMatch = raw.match(
                /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/
            );

            if (brDateMatch) {
                month = parseInt(brDateMatch[2], 10);
                year = normalizeYear(parseInt(brDateMatch[3], 10));
            }
        }

        // Formato mes numerico e ano: 01/2026 ou 1-26
        if (month === null || year === null) {
            const monthYearMatch = raw.match(
                /^(\d{1,2})[\/.-](\d{2}|\d{4})$/
            );

            if (monthYearMatch) {
                month = parseInt(monthYearMatch[1], 10);
                year = normalizeYear(parseInt(monthYearMatch[2], 10));
            }
        }

        // Formato mes por extenso ou abreviado: jan/26, janeiro/2026, fev-26
        if (month === null || year === null) {
            const textMonthMatch = raw.match(
                /^([A-Za-zA-u]+)\s*[\/.-]\s*(\d{2}|\d{4})$/i
            );

            if (textMonthMatch) {
                month = getMonthNumber(textMonthMatch[1]);
                year = normalizeYear(parseInt(textMonthMatch[2], 10));
            }
        }

        // Ultima tentativa para textos reconhecidos pelo JavaScript, como "May 1 2026".
        if (month === null || year === null) {
            const parsedDate = new Date(raw);

            if (!isNaN(parsedDate.getTime()) && onlyDigits(raw).length >= 4) {
                month = parsedDate.getMonth() + 1;
                year = parsedDate.getFullYear();
            }
        }

        // Caso o valor nao seja reconhecido como uma data, mantem o proprio valor como coluna.
        if (month === null || year === null || month < 1 || month > 12) {
            return {
                key: `text:${raw}`,
                label: raw,
                sortValue: Number.MAX_SAFE_INTEGER
            };
        }
    }

    if (month === null || year === null || month < 1 || month > 12) {
        return null;
    }

    const shortYear = String(year).slice(-2);
    const key = `${year}-${String(month).padStart(2, "0")}`;

    return {
        key,
        label: `${monthNames[month]}/${shortYear}`,
        sortValue: year * 100 + month
    };
}

/**
 * Formata valores numericos para exibicao.
 */
function formatNumber(value: number): string {
    if (!isFinite(value)) {
        return "";
    }

    return new Intl.NumberFormat("pt-BR", {
        maximumFractionDigits: 6
    }).format(value);
}

/**
 * Verifica se uma celula possui valor visivel.
 *
 * Zero e considerado um valor valido.
 */
function hasVisibleValue(value: unknown): boolean {
    return value !== null && value !== undefined && value !== "";
}

/**
 * Linha utilizada no detalhamento.
 */
interface DetailRow {
    values: Map<number, unknown>;
}

/**
 * Linha da matriz ja agrupada pelas colunas informativas.
 */
interface PivotRow {
    info: Map<number, unknown>;
    periodValues: Map<string, unknown>;
    details: Map<string, DetailRow[]>;
}

/**
 * Cores resolvidas do popup, aplicadas dinamicamente ao construir
 * o conteudo do detalhamento.
 */
interface PopupColors {
    background: string;
    fontColor: string;
    headerBackground: string;
    rowAltColor: string;
    borderColor: string;
}

/**
 * Visual principal.
 */
export class Visual implements IVisual {
    private formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;

    private root: HTMLElement;
    private titleElement: HTMLElement;
    private scrollWrap: HTMLElement;
    private overlay: HTMLElement;
    private modal: HTMLElement;
    private modalTitle: HTMLElement;
    private modalClose: HTMLElement;
    private modalBody: HTMLElement;

    private currentPivotRows: PivotRow[] = [];
    private currentPeriods: PeriodInfo[] = [];
    private currentDetailCols: DataViewMetadataColumn[] = [];
    private currentInfoCols: DataViewMetadataColumn[] = [];

    private popupColors: PopupColors = {
        background: "#2E4153",
        fontColor: "#FFFFFF",
        headerBackground: "#3F5B70",
        rowAltColor: "#364D60",
        borderColor: "#4E6E85"
    };

    private exportHeaders: string[] = [];
    private exportRows: string[][] = [];

    /**
     * Cria os elementos permanentes do visual.
     */
    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();

        injectStyle("fperiod-style", VISUAL_CSS);

        this.root = document.createElement("div");
        this.root.className = "fperiod-wrap";
        options.element.appendChild(this.root);

        this.titleElement = document.createElement("div");
        this.titleElement.className = "fperiod-title";
        this.root.appendChild(this.titleElement);

        this.scrollWrap = document.createElement("div");
        this.scrollWrap.className = "fperiod-scroll";
        this.root.appendChild(this.scrollWrap);

        this.scrollWrap.addEventListener(
            "click",
            (event: MouseEvent) => this.handleTableClick(event)
        );

        this.overlay = document.createElement("div");
        this.overlay.className = "fperiod-overlay";
        this.overlay.addEventListener("click", () => this.closeModal());
        this.root.appendChild(this.overlay);

        this.modal = document.createElement("div");
        this.modal.className = "fperiod-modal";

        const modalHeader = document.createElement("div");
        modalHeader.className = "fperiod-modal-header";

        this.modalTitle = document.createElement("span");
        this.modalTitle.className = "fperiod-modal-title";
        modalHeader.appendChild(this.modalTitle);

        this.modalClose = document.createElement("button");
        this.modalClose.className = "fperiod-modal-close";
        (this.modalClose as HTMLButtonElement).type = "button";
        this.modalClose.textContent = "X";
        this.modalClose.setAttribute("aria-label", "Fechar detalhamento");
        this.modalClose.addEventListener("click", () => this.closeModal());

        modalHeader.appendChild(this.modalClose);
        this.modal.appendChild(modalHeader);

        this.modalBody = document.createElement("div");
        this.modal.appendChild(this.modalBody);

        this.root.appendChild(this.modal);

        document.addEventListener("keydown", (event: KeyboardEvent) => {
            if (event.key === "Escape") {
                this.closeModal();
            }
        });
    }

    /**
     * Retorna as configuracoes para o painel de formatacao.
     */
    public getFormattingModel(): FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(
            this.formattingSettings
        );
    }

    /**
     * Atualiza o visual quando os dados, filtros, tamanho ou
     * configuracoes forem alterados.
     */
    public update(options: VisualUpdateOptions): void {
        const dataView = options.dataViews?.[0];

        this.formattingSettings =
            this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView
            );

        this.applyFormatting();

        if (!dataView?.table) {
            this.showEmptyMessage(
                "Nenhum dado disponivel. Adicione campos ao visual."
            );
            return;
        }

        const table = dataView.table;

        const firstInfoIndex = table.columns.findIndex(
            column => Boolean(column.roles?.["info"])
        );

        const sortedRows = [...(table.rows ?? [])];

        if (firstInfoIndex >= 0) {
            sortedRows.sort(
                (firstRow: DataViewTableRow, secondRow: DataViewTableRow) => {
                    const firstValue = String(
                        firstRow[firstInfoIndex] ?? ""
                    ).toLocaleLowerCase("pt-BR");

                    const secondValue = String(
                        secondRow[firstInfoIndex] ?? ""
                    ).toLocaleLowerCase("pt-BR");

                    return firstValue.localeCompare(secondValue, "pt-BR");
                }
            );
        }

        const sortedTable: DataViewTable = {
            columns: table.columns,
            rows: sortedRows,
            totals: table.totals
        };

        this.render(sortedTable);
    }

    /**
     * Aplica as configuracoes escolhidas no painel do Power BI aos
     * elementos permanentes (titulo, popup e overlay).
     */
    private applyFormatting(): void {
        const titleSettings = this.formattingSettings.titleCard;
        const popupSettings = this.formattingSettings.popupCard;

        // Titulo.
        const showTitle = titleSettings.show.value;
        const titleText = titleSettings.text.value?.trim() ?? "";

        this.titleElement.textContent = titleText;

        if (showTitle && titleText.length > 0) {
            this.titleElement.classList.remove("fperiod-title-hidden");
        } else {
            this.titleElement.classList.add("fperiod-title-hidden");
        }

        this.titleElement.style.fontSize = `${titleSettings.fontSize.value}px`;
        this.titleElement.style.color = titleSettings.fontColor.value.value;
        this.titleElement.style.backgroundColor =
            titleSettings.backgroundColor.value.value;

        // Popup: guarda as cores para uso ao montar o conteudo.
        this.popupColors = {
            background: popupSettings.backgroundColor.value.value,
            fontColor: popupSettings.fontColor.value.value,
            headerBackground: popupSettings.headerBackgroundColor.value.value,
            rowAltColor: popupSettings.rowAltColor.value.value,
            borderColor: popupSettings.borderColor.value.value
        };

        // Overlay escurecido a partir da cor escolhida, com transparencia.
        this.overlay.style.backgroundColor = hexToRgba(
            popupSettings.overlayColor.value.value,
            0.45
        );

        // Estilos fixos do container do modal.
        this.modal.style.background = this.popupColors.background;
        this.modal.style.color = this.popupColors.fontColor;
        this.modal.style.border = `1px solid ${this.popupColors.borderColor}`;

        this.modalTitle.style.color = this.popupColors.fontColor;

        this.modalClose.style.background = this.popupColors.headerBackground;
        this.modalClose.style.color = this.popupColors.fontColor;
    }

    /**
     * Limpa os dados internos e mostra uma mensagem.
     */
    private showEmptyMessage(message: string): void {
        this.scrollWrap.innerHTML =
            `<div class="fperiod-empty">${escapeHtml(message)}</div>`;

        this.exportHeaders = [];
        this.exportRows = [];
        this.currentPivotRows = [];
        this.currentPeriods = [];
        this.currentDetailCols = [];
        this.currentInfoCols = [];
    }

    /**
     * Formata um valor com base nos metadados da coluna.
     */
    private formatValue(
        column: DataViewMetadataColumn,
        value: unknown
    ): string {
        if (value === null || value === undefined) {
            return "";
        }

        if (column.type?.dateTime) {
            const date = new Date(value as string | number | Date);

            if (!isNaN(date.getTime())) {
                const day = String(date.getDate()).padStart(2, "0");
                const month = String(date.getMonth() + 1).padStart(2, "0");
                const year = date.getFullYear();

                return `${day}/${month}/${year}`;
            }
        }

        if (typeof value === "number") {
            return formatNumber(value);
        }

        return String(value);
    }

    /**
     * Define o valor que sera mantido quando existirem varias
     * linhas para a mesma combinacao de dados e periodo.
     *
     * Quando os valores forem numericos, mantem o maior valor,
     * preservando o comportamento do componente original.
     */
    private mergePeriodValue(currentValue: unknown, newValue: unknown): unknown {
        if (!hasVisibleValue(currentValue)) {
            return newValue;
        }

        if (!hasVisibleValue(newValue)) {
            return currentValue;
        }

        const currentNumber =
            typeof currentValue === "number"
                ? currentValue
                : Number(currentValue);

        const newNumber =
            typeof newValue === "number" ? newValue : Number(newValue);

        if (!isNaN(currentNumber) && !isNaN(newNumber)) {
            return Math.max(currentNumber, newNumber);
        }

        return currentValue;
    }

    /**
     * Cria a matriz agrupada e renderiza o HTML.
     */
    private render(table: DataViewTable): void {
        const columns = table.columns;
        const rows = table.rows ?? [];

        const infoColumns: DataViewMetadataColumn[] = [];
        const periodColumns: DataViewMetadataColumn[] = [];
        const valueColumns: DataViewMetadataColumn[] = [];
        const detailColumns: DataViewMetadataColumn[] = [];

        columns.forEach(column => {
            if (column.roles?.["info"]) {
                infoColumns.push(column);
            } else if (column.roles?.["period"]) {
                periodColumns.push(column);
            } else if (column.roles?.["value"]) {
                valueColumns.push(column);
            } else if (column.roles?.["detail"]) {
                detailColumns.push(column);
            }
        });

        const periodIndex = periodColumns[0]?.index;
        const valueIndex = valueColumns[0]?.index;

        if (periodIndex === undefined || valueIndex === undefined) {
            this.showEmptyMessage("Configure os campos Data e Valor.");
            return;
        }

        const pivotMap = new Map<string, PivotRow>();
        const periodMap = new Map<string, PeriodInfo>();
        const detailSeenKeys = new Set<string>();

        rows.forEach((row: DataViewTableRow) => {
            const period = createPeriodInfo(row[periodIndex]);

            if (!period) {
                return;
            }

            if (!periodMap.has(period.key)) {
                periodMap.set(period.key, period);
            }

            // O separador inclui o indice e o tamanho do texto para reduzir
            // a possibilidade de duas combinacoes produzirem a mesma chave.
            const infoKey = infoColumns
                .map(column => {
                    const rawValue = String(row[column.index!] ?? "");
                    return `${column.index}:${rawValue.length}:${rawValue}`;
                })
                .join("||");

            if (!pivotMap.has(infoKey)) {
                const infoMap = new Map<number, unknown>();

                infoColumns.forEach(column => {
                    infoMap.set(column.index!, row[column.index!]);
                });

                pivotMap.set(infoKey, {
                    info: infoMap,
                    periodValues: new Map<string, unknown>(),
                    details: new Map<string, DetailRow[]>()
                });
            }

            const pivotRow = pivotMap.get(infoKey)!;
            const currentValue = pivotRow.periodValues.get(period.key);

            const mergedValue = this.mergePeriodValue(
                currentValue,
                row[valueIndex]
            );

            pivotRow.periodValues.set(period.key, mergedValue);

            if (detailColumns.length > 0) {
                const detailMap = new Map<number, unknown>();

                detailColumns.forEach(column => {
                    detailMap.set(column.index!, row[column.index!]);
                });

                const detailKey = [
                    infoKey,
                    period.key,
                    ...detailColumns.map(column =>
                        String(row[column.index!] ?? "")
                    )
                ].join("||");

                if (!detailSeenKeys.has(detailKey)) {
                    detailSeenKeys.add(detailKey);

                    const existingDetails =
                        pivotRow.details.get(period.key) ?? [];

                    existingDetails.push({ values: detailMap });

                    pivotRow.details.set(period.key, existingDetails);
                }
            }
        });

        const periods = Array.from(periodMap.values()).sort(
            (firstPeriod, secondPeriod) => {
                if (firstPeriod.sortValue !== secondPeriod.sortValue) {
                    return firstPeriod.sortValue - secondPeriod.sortValue;
                }

                return firstPeriod.label.localeCompare(
                    secondPeriod.label,
                    "pt-BR"
                );
            }
        );

        const pivotRows = Array.from(pivotMap.values()).sort(
            (firstRow, secondRow) => {
                const firstInfoColumn = infoColumns[0];

                if (!firstInfoColumn) {
                    return 0;
                }

                const firstValue = String(
                    firstRow.info.get(firstInfoColumn.index!) ?? ""
                ).toLocaleLowerCase("pt-BR");

                const secondValue = String(
                    secondRow.info.get(firstInfoColumn.index!) ?? ""
                ).toLocaleLowerCase("pt-BR");

                return firstValue.localeCompare(secondValue, "pt-BR");
            }
        );

        this.currentPivotRows = pivotRows;
        this.currentPeriods = periods;
        this.currentDetailCols = detailColumns;
        this.currentInfoCols = infoColumns;

        this.buildTable(
            pivotRows,
            periods,
            infoColumns,
            valueColumns[0],
            detailColumns
        );

        this.prepareCsv(pivotRows, periods, infoColumns, valueColumns[0]);
    }

    /**
     * Monta o HTML da tabela.
     */
    private buildTable(
        pivotRows: PivotRow[],
        periods: PeriodInfo[],
        infoColumns: DataViewMetadataColumn[],
        valueColumn: DataViewMetadataColumn,
        detailColumns: DataViewMetadataColumn[]
    ): void {
        const tableSettings = this.formattingSettings.tableCard;
        const totalsSettings = this.formattingSettings.totalsCard;

        const headerFontColor = tableSettings.headerFontColor.value.value;
        const headerBackground = tableSettings.headerBackgroundColor.value.value;
        const groupBackground =
            tableSettings.groupHeaderBackgroundColor.value.value;
        const oddRowColor = tableSettings.oddRowColor.value.value;
        const oddRowFontColor = tableSettings.oddRowFontColor.value.value;
        const evenRowColor = tableSettings.evenRowColor.value.value;
        const evenRowFontColor = tableSettings.evenRowFontColor.value.value;
        const gridColor = tableSettings.gridColor.value.value;
        const fontSize = tableSettings.fontSize.value;

        // Cores dos totais.
        const totalRowBg = totalsSettings.rowBackgroundColor.value.value;
        const totalRowColor = totalsSettings.rowFontColor.value.value;
        const totalColBg = totalsSettings.columnBackgroundColor.value.value;
        const totalColColor = totalsSettings.columnFontColor.value.value;
        const cornerBg = totalsSettings.cornerBackgroundColor.value.value;
        const cornerColor = totalsSettings.cornerFontColor.value.value;

        // Total de colunas: informativas + meses + coluna Total Geral.
        const totalColumns = infoColumns.length + periods.length + 1;

        // Acumuladores para os totais.
        const columnTotals = new Map<string, number>();
        let grandTotal = 0;
        let hasAnyNumber = false;

        periods.forEach(period => columnTotals.set(period.key, 0));

        const parts: string[] = [];

        parts.push(
            `<table class="fperiod-table" style="font-size:${fontSize}px;">`
        );

        parts.push("<thead>");

        // Primeira linha: Dados Informativos, Meses e Total Geral.
        parts.push("<tr>");

        if (infoColumns.length > 0) {
            parts.push(
                `<th class="fperiod-th-group" ` +
                `colspan="${infoColumns.length}" ` +
                `style="background:${groupBackground};` +
                `color:${headerFontColor};` +
                `border:1px solid ${gridColor};">` +
                `Dados Informativos` +
                `</th>`
            );
        }

        if (periods.length > 0) {
            parts.push(
                `<th class="fperiod-th-group" ` +
                `colspan="${periods.length}" ` +
                `style="background:${groupBackground};` +
                `color:${headerFontColor};` +
                `border:1px solid ${gridColor};">` +
                `Meses` +
                `</th>`
            );
        }

        // Cabecalho de grupo da coluna Total Geral.
        parts.push(
            `<th class="fperiod-th-group" ` +
            `colspan="1" ` +
            `style="background:${totalColBg};` +
            `color:${totalColColor};` +
            `border:1px solid ${gridColor};">` +
            `Total Geral` +
            `</th>`
        );

        parts.push("</tr>");

        // Segunda linha: nomes das colunas informativas, nomes dos meses e "Total".
        parts.push("<tr>");

        infoColumns.forEach(column => {
            parts.push(
                `<th class="fperiod-th-info" ` +
                `style="background:${headerBackground};` +
                `color:${headerFontColor};` +
                `border:1px solid ${gridColor};">` +
                `${escapeHtml(column.displayName)}` +
                `</th>`
            );
        });

        periods.forEach(period => {
            parts.push(
                `<th class="fperiod-th-period" ` +
                `style="background:${headerBackground};` +
                `color:${headerFontColor};` +
                `border:1px solid ${gridColor};">` +
                `${escapeHtml(period.label)}` +
                `</th>`
            );
        });

        // Sub-cabecalho da coluna Total Geral.
        parts.push(
            `<th class="fperiod-th-total" ` +
            `style="background:${totalColBg};` +
            `color:${totalColColor};` +
            `border:1px solid ${gridColor};">` +
            `Total` +
            `</th>`
        );

        parts.push("</tr>");
        parts.push("</thead>");
        parts.push("<tbody>");

        if (pivotRows.length === 0) {
            parts.push(
                `<tr>` +
                `<td colspan="${Math.max(totalColumns, 1)}" ` +
                `class="fperiod-empty" ` +
                `style="border:1px solid ${gridColor};">` +
                `Sem registros para o periodo selecionado.` +
                `</td>` +
                `</tr>`
            );
        } else {
            pivotRows.forEach((pivotRow, rowIndex) => {
                const isOdd = rowIndex % 2 === 0;
                const rowColor = isOdd ? oddRowColor : evenRowColor;
                const rowFontColor = isOdd ? oddRowFontColor : evenRowFontColor;

                parts.push("<tr>");

                // Colunas informativas primeiro.
                infoColumns.forEach(column => {
                    const value = pivotRow.info.get(column.index!);

                    parts.push(
                        `<td class="fperiod-td-info" ` +
                        `style="background:${rowColor};` +
                        `color:${rowFontColor};` +
                        `border:1px solid ${gridColor};">` +
                        `${escapeHtml(this.formatValue(column, value))}` +
                        `</td>`
                    );
                });

                // Meses depois das colunas informativas.
                let rowTotal = 0;
                let rowHasNumber = false;

                periods.forEach((period, periodIndex) => {
                    const rawValue = pivotRow.periodValues.get(period.key);
                    const visible = hasVisibleValue(rawValue);
                    const details = pivotRow.details.get(period.key) ?? [];

                    const numeric = toNumber(rawValue);

                    if (numeric !== null) {
                        rowTotal += numeric;
                        rowHasNumber = true;
                        hasAnyNumber = true;
                        columnTotals.set(
                            period.key,
                            (columnTotals.get(period.key) ?? 0) + numeric
                        );
                    }

                    const clickable =
                        detailColumns.length > 0 && details.length > 0;

                    const cssClickable = clickable ? " clickable" : "";

                    const dataAttributes = clickable
                        ? ` data-row="${rowIndex}" data-col="${periodIndex}"`
                        : "";

                    const displayValue = visible
                        ? this.formatValue(valueColumn, rawValue)
                        : "";

                    parts.push(
                        `<td class="fperiod-td-period${cssClickable}"` +
                        `${dataAttributes} ` +
                        `style="background:${rowColor};` +
                        `color:${rowFontColor};` +
                        `border:1px solid ${gridColor};">` +
                        `${escapeHtml(displayValue)}` +
                        `</td>`
                    );
                });

                // Coluna Total Geral da linha.
                grandTotal += rowTotal;

                const rowTotalDisplay = rowHasNumber
                    ? formatNumber(rowTotal)
                    : "";

                parts.push(
                    `<td class="fperiod-td-total" ` +
                    `style="background:${totalColBg};` +
                    `color:${totalColColor};` +
                    `border:1px solid ${gridColor};">` +
                    `${escapeHtml(rowTotalDisplay)}` +
                    `</td>`
                );

                parts.push("</tr>");
            });
        }

        parts.push("</tbody>");

        // Rodape fixo (Total Mensal).
        if (pivotRows.length > 0) {
            parts.push("<tfoot>");
            parts.push("<tr>");

            // Rotulo "Total Mensal" ocupando as colunas informativas.
            const labelColspan = Math.max(infoColumns.length, 1);

            parts.push(
                `<td class="fperiod-tfoot-cell fperiod-tfoot-label fperiod-tfoot-corner" ` +
                `colspan="${labelColspan}" ` +
                `style="background:${totalRowBg};` +
                `color:${totalRowColor};` +
                `border:1px solid ${gridColor};">` +
                `Total Mensal` +
                `</td>`
            );

            // Total de cada mes.
            periods.forEach(period => {
                const total = columnTotals.get(period.key) ?? 0;
                const display = hasAnyNumber ? formatNumber(total) : "";

                parts.push(
                    `<td class="fperiod-tfoot-cell" ` +
                    `style="background:${totalRowBg};` +
                    `color:${totalRowColor};` +
                    `border:1px solid ${gridColor};">` +
                    `${escapeHtml(display)}` +
                    `</td>`
                );
            });

            // Total geral (canto inferior direito).
            const grandDisplay = hasAnyNumber ? formatNumber(grandTotal) : "";

            parts.push(
                `<td class="fperiod-tfoot-cell fperiod-tfoot-corner" ` +
                `style="background:${cornerBg};` +
                `color:${cornerColor};` +
                `border:1px solid ${gridColor};">` +
                `${escapeHtml(grandDisplay)}` +
                `</td>`
            );

            parts.push("</tr>");
            parts.push("</tfoot>");
        }

        parts.push("</table>");

        this.scrollWrap.innerHTML = parts.join("");

        // Ajusta a posicao sticky da segunda linha do cabecalho com base
        // na altura efetiva da primeira linha.
        window.requestAnimationFrame(() => {
            const groupHeader = this.scrollWrap.querySelector<HTMLElement>(
                ".fperiod-th-group"
            );

            if (!groupHeader) {
                return;
            }

            const groupHeight = groupHeader.getBoundingClientRect().height;

            this.scrollWrap
                .querySelectorAll<HTMLElement>(
                    ".fperiod-th-info, .fperiod-th-period, .fperiod-th-total"
                )
                .forEach(element => {
                    element.style.top = `${groupHeight}px`;
                });
        });
    }

    /**
     * Trata o clique nas celulas mensais.
     */
    private handleTableClick(event: MouseEvent): void {
        const target = event.target as HTMLElement;

        const cell = target.closest(
            "td[data-row][data-col]"
        ) as HTMLTableCellElement | null;

        if (!cell) {
            return;
        }

        const rowIndex = parseInt(cell.getAttribute("data-row") ?? "-1", 10);
        const periodIndex = parseInt(cell.getAttribute("data-col") ?? "-1", 10);

        if (rowIndex < 0 || periodIndex < 0) {
            return;
        }

        const pivotRow = this.currentPivotRows[rowIndex];
        const period = this.currentPeriods[periodIndex];

        if (!pivotRow || !period || this.currentDetailCols.length === 0) {
            return;
        }

        const detailRows = pivotRow.details.get(period.key) ?? [];
        const firstInfoColumn = this.currentInfoCols[0];

        const firstInfoValue = firstInfoColumn
            ? String(pivotRow.info.get(firstInfoColumn.index!) ?? "")
            : "";

        const formattedDetails: DetailRow[] = detailRows.map(detailRow => {
            const formattedMap = new Map<number, unknown>();

            this.currentDetailCols.forEach(column => {
                const rawValue = detailRow.values.get(column.index!);
                formattedMap.set(
                    column.index!,
                    this.formatValue(column, rawValue)
                );
            });

            return { values: formattedMap };
        });

        const modalTitle = firstInfoValue
            ? `${firstInfoValue} | ${period.label}`
            : period.label;

        this.openModal(modalTitle, formattedDetails, this.currentDetailCols);
    }

    /**
     * Abre o popup de detalhamento, aplicando as cores configuradas.
     */
    private openModal(
        title: string,
        detailRows: DetailRow[],
        detailColumns: DataViewMetadataColumn[]
    ): void {
        this.modalTitle.textContent = title;

        const colors = this.popupColors;

        if (detailRows.length === 0) {
            this.modalBody.innerHTML =
                `<p style="color:${colors.fontColor};` +
                `text-align:center;padding:20px;">` +
                `Sem registros de detalhe para este periodo.` +
                `</p>`;
        } else {
            const parts: string[] = [];

            parts.push(`<table class="fperiod-modal-table">`);
            parts.push("<thead><tr>");

            detailColumns.forEach(column => {
                parts.push(
                    `<th style="background:${colors.headerBackground};` +
                    `color:${colors.fontColor};` +
                    `border:1px solid ${colors.borderColor};">` +
                    `${escapeHtml(column.displayName)}` +
                    `</th>`
                );
            });

            parts.push("</tr></thead>");
            parts.push("<tbody>");

            detailRows.forEach((detailRow, rowIndex) => {
                const rowBg =
                    rowIndex % 2 === 0
                        ? colors.background
                        : colors.rowAltColor;

                parts.push("<tr>");

                detailColumns.forEach(column => {
                    parts.push(
                        `<td style="background:${rowBg};` +
                        `color:${colors.fontColor};` +
                        `border:1px solid ${colors.borderColor};">` +
                        `${escapeHtml(detailRow.values.get(column.index!))}` +
                        `</td>`
                    );
                });

                parts.push("</tr>");
            });

            parts.push("</tbody>");
            parts.push("</table>");

            this.modalBody.innerHTML = parts.join("");
        }

        this.overlay.style.display = "block";
        this.modal.style.display = "block";
    }

    /**
     * Fecha o popup de detalhamento.
     */
    private closeModal(): void {
        this.overlay.style.display = "none";
        this.modal.style.display = "none";
    }

    /**
     * Prepara os dados na mesma ordem apresentada na tabela,
     * incluindo a coluna Total Geral e a linha Total Mensal.
     */
    private prepareCsv(
        pivotRows: PivotRow[],
        periods: PeriodInfo[],
        infoColumns: DataViewMetadataColumn[],
        valueColumn: DataViewMetadataColumn
    ): void {
        this.exportHeaders = [
            ...infoColumns.map(column => column.displayName),
            ...periods.map(period => period.label),
            "Total Geral"
        ];

        const columnTotals = new Map<string, number>();
        let grandTotal = 0;

        periods.forEach(period => columnTotals.set(period.key, 0));

        this.exportRows = pivotRows.map(pivotRow => {
            let rowTotal = 0;

            const periodCells = periods.map(period => {
                const value = pivotRow.periodValues.get(period.key);
                const numeric = toNumber(value);

                if (numeric !== null) {
                    rowTotal += numeric;
                    columnTotals.set(
                        period.key,
                        (columnTotals.get(period.key) ?? 0) + numeric
                    );
                }

                return hasVisibleValue(value)
                    ? this.formatValue(valueColumn, value)
                    : "";
            });

            grandTotal += rowTotal;

            return [
                ...infoColumns.map(column => {
                    const value = pivotRow.info.get(column.index!);
                    return this.formatValue(column, value);
                }),
                ...periodCells,
                formatNumber(rowTotal)
            ];
        });

        // Linha final com os totais mensais.
        if (pivotRows.length > 0) {
            const totalRow: string[] = [];

            infoColumns.forEach((_column, index) => {
                totalRow.push(index === 0 ? "Total Mensal" : "");
            });

            if (infoColumns.length === 0) {
                totalRow.push("Total Mensal");
            }

            periods.forEach(period => {
                totalRow.push(formatNumber(columnTotals.get(period.key) ?? 0));
            });

            totalRow.push(formatNumber(grandTotal));

            this.exportRows.push(totalRow);
        }
    }

    /**
     * Cria o conteudo CSV com separador ponto e virgula.
     *
     * O BOM e incluido para preservar acentuacao ao abrir no Excel.
     */
    private buildCsvContent(withBom: boolean = true): string {
        const escapeCsv = (value: string): string =>
            `"${value.replace(/"/g, "\"\"")}"`;

        const lines: string[] = [];

        lines.push(this.exportHeaders.map(escapeCsv).join(";"));

        this.exportRows.forEach(row => {
            lines.push(row.map(escapeCsv).join(";"));
        });

        const csv = lines.join("\r\n");

        return withBom ? `\uFEFF${csv}` : csv;
    }
}
