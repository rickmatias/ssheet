import Model from '@mdev-js/model';
import { differenceInSeconds, isDate, isSameMinute } from 'date-fns';

/**
 * @typedef {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheet
 * @typedef {GoogleAppsScript.Spreadsheet.Sheet} Sheet
 * @typedef {GoogleAppsScript.Spreadsheet.Range} Range
 * @typedef {GoogleAppsScript.Drive.File} File
 */

 /**
 * @typedef {object} ConstructorParams
 * @property {string|string[]} [primaryKey] - A chave primária ou array de chaves primárias.
 * @property {string} sheetName - O nome da planilha.
 * @property {string} ssId - O ID da planilha.
 */

/**
 * @typedef {Object} RowObject
 * @property {number} [rowNumber] - O número da linha, opcional.
 * @property {*} [property1] - Uma propriedade adicional de qualquer tipo.
 * @property {*} [property2] - Outra propriedade adicional de qualquer tipo.
 * ...
 */

/**
 * @param {any[]} arrayA
 * @param {any[]} arrayB
 * @returns {any[]}
 */
const mergeThreeDimensionalArrays = (arrayA, arrayB) => {
  const mergeTwoDimentionalArrays = (arrayA, arrayB) => {
    const biggestArray = arrayA.length >= arrayB.length ? arrayA : arrayB;
    const smallestArray = arrayA.length >= arrayB.length ? arrayB : arrayA;

    return biggestArray.map((value, index) => value || smallestArray[index]);
  };

  if(arrayA.length !== arrayB.length)
    throw new Error('Arrays A and B must have the same size');

  return arrayA.map((arrayRowA, index) => {
    const arrayRowB = arrayB[index];
    if(arrayRowA.length !== arrayRowB.length){
      throw new Error('The rows from array A and B with index ${rowIndex} has different sizes');
    }
    return mergeTwoDimentionalArrays(arrayRowA, arrayRowB);
  });
};

/**
 * @class SSheet
 */
export default class SSheet {
  /**
   * @param {ConstructorParams} params
   */
  constructor({ primaryKey, sheetName, ssId }) {
    /**
     * @type {string[]}
     * @private
     */
    this._primaryKey = primaryKey
      ? Array.isArray(primaryKey)
        ? primaryKey
        : [primaryKey]
      : ['rowNumber'];

    /**
     * @type {string}
     * @private
     */
    this._sheetName = sheetName;

    /**
     * @type {string}
     * @private
     */
    this._ssId = ssId;

    /**
     * @type {object}
     * @private
     */
    this._columnMap = undefined;

    /**
     * @type {string[]}
     * @private
     */
    this._columnsWithMapFormula = undefined;

    /**
     * @type {any[][]}
     * @private
     */
    this._data = [];

    /**
     * @type {Model}
     * @private
     */
    this._entity = undefined;

    /**
     * @type {number}
     * @private
     */
    this._headerRow = undefined;

    /**
     * @type {boolean}
     * @private
     */
    this._inTransaction = false;

    /**
     * @type {number}
     * @private
     */
    this._lastColumn = undefined;

    /**
     * @type {Date}
     * @private
     */
    this._lastRefreshInCache = undefined;

    /**
     * @type {number}
     * @private
     */
    this._lastRow = undefined;

    /**
     * @type {Object.<string, number>}
     * @private
     */
    this._mapByPrimaryKey = {};

    /**
     * @type {object}
     * @private
     */
    this._parameterMap = undefined;

    /**
     * @type {object}
     * @private
     */
    this._rangesA1 = undefined;

    /**
     * @type {object}
     * @private
     */
    this._rangesWithFormulas = undefined;

    /**
     * @type {string[]}
     * @private
     */
    this._readOnlyKeys = [];

    /**
     * @type {Sheet}
     * @private
     */
    this._sheet = undefined;

    /**
     * @type {Spreadsheet}
     * @private
     */
    this._spreadsheet = undefined;

    /**
     * @type {boolean}
     * @private
     */
    this._keepSheetHidden = false;
  }

  /**
   * Returns the child class name.
   * @returns {string}
   * @readonly
   */
  get className() {
    return this.constructor.name;
  }

  /**
   * @returns {object}
   */
  get columnMap() {
    return this._columnMap;
  }

  /**
   * @param {object} columnMap
   * @returns {void}
   */
  set columnMap(columnMap) {
    this._columnMap = columnMap;
  }

  get columnsWithMapFormula() {
    if(!this._columnsWithMapFormula){
      const range = this.sheet.getRange(this.headerRow, 1, 1, this.lastColumn);
      const formulas = mergeThreeDimensionalArrays(
        range.getFormulas(),
        range.getFormulasR1C1()
      )[0];

      this._columnsWithMapFormula = formulas
        .reduce((/**@type {any[]}*/ arr, /**@type{string}*/ value, /**@type{number}*/ index) => {
          if(value && value.match(/^\=MAP.*/i)){
            arr.push(this.getColNameByColNumber(index + 1));
          }
          return arr;
        },[]);
    }
    return this._columnsWithMapFormula;
  }

  get data() {
    return this._data;
  }

  /**
   * @returns {object}
   */
  get entity() {
    return this._entity;
  }

  /**
   * @param {Model} entity
   * @returns {void}
   */
  set entity(entity) {
    this._entity = entity;
  }

  /**
   * @returns {number}
   */
  get headerRow() {
    if (!this._headerRow) this._headerRow = 1;

    return this._headerRow;
  }

  /**
   * @returns {any[]}
   * @readonly
   */
  get headerRowData() {
    const {data, headerRow} = this;
    return data[headerRow-1] || this.getRowData(headerRow);
  }

  /**
   * @returns {boolean}
   */
  get keepSheetHidden() {
    return this._keepSheetHidden;
  }

  /**
   * @param {boolean} keepSheetHidden
   * @returns {void}
   */
  set keepSheetHidden(keepSheetHidden) {
    this._keepSheetHidden = keepSheetHidden;
  }

  /**
   * @returns {number}
   * @readonly
   */
  get lastColumn() {
    if (!this._lastColumn) {
      this._lastColumn = this.sheet.getLastColumn();
    }
    return this._lastColumn;
  }

  get lastRefreshInCache(){
    return this._lastRefreshInCache;
  }

  get mapByPrimaryKey() {
    return this._mapByPrimaryKey;
  }

  /**
   * @returns {number}
   * @readonly
   */
  get lastRow() {
    if (!this._lastRow) {
      SpreadsheetApp.flush();
      this._lastRow = this.sheet.getLastRow();
    }
    return this._lastRow;
  }

  /**
   * @returns {string[]} A lista de colunas que compõem a chave primária da planilha.
   * @readonly
   */
  get primaryKey() {
    return this._primaryKey;
  }

  /**
   * @returns {object}
   */
  get rangesA1() {
    return this._rangesA1;
  }

  /**
   * @param {object} rangesA1
   * @returns {void}
   */
  set rangesA1(rangesA1) {
    this._rangesA1 = rangesA1;
  }

  /**
   * @returns {object}
   */
  get rangesWithFormulas() {
    return this._rangesWithFormulas;
  }

  /**
   * @param {object} rangesWithFormulas
   * @returns {void}
   */
  set rangesWithFormulas(rangesWithFormulas) {
    this._rangesWithFormulas = rangesWithFormulas;
  }

  /**
   * @returns {string[]}
   */
  get readOnlyKeys() {
    return this._readOnlyKeys;
  }

  /**
   * @param {string[]} readOnlyKeys
   * @returns {void}
   */
  set readOnlyKeys(readOnlyKeys) {
    this._readOnlyKeys = [...readOnlyKeys];
  }

  /**
   * @return {Sheet} A planilha associada ao objeto SSheet.
   */
  get sheet() {
    if (!this._sheet) {
      this._sheet = SSheet.safelyGetSheetByName(
        this.spreadsheet,
        this.sheetName
      );
    }
    return this._sheet;
  }

  /**
   * @returns {number}
   * @readonly
   */
  get sheetId() {
    return this.sheet.getSheetId();
  }

  /**
   * @returns {string}
   * @readonly
   */
  get sheetName() {
    return this._sheetName;
  }

  /**
   * @returns {Spreadsheet}
   * @readonly
   */
  get spreadsheet() {
    if (!this._spreadsheet) {
      this._spreadsheet = SSheet.safelyOpenSpreadsheetById(this.ssId);
    }
    return this._spreadsheet;
  }

  /**
   * @returns {string}
   * @readonly
   */
  get ssId() {
    return this._ssId;
  }

  /**
   * @param {number} headerRow
   * @returns {void}
   * @throws {string} Se o parâmetro "headerRow" não for um número inteiro maior ou igual a 1.
   */
  set headerRow(headerRow) {
    if (!(headerRow && Number.isInteger(headerRow) && headerRow >= 1)) {
      throw 'Ops! O parâmetro "headerRow" deve ser um número inteiro maior ou igual a 1.';
    }
    this._headerRow = headerRow;
  }

  /**
   * Inicia uma transação na planilha, para que as alterações sejam feitas em lote.
   * @returns {void}
   */
  beginTransaction() {
    this._inTransaction = true;
  }

  /**
   * @param {string} key
   * @returns {any}
   */
  get(key) {
    return this.getRangeVal(this.rangesA1[key]);
  }

  /**
   * @typedef {Object} CallbackDTO
   * @property {string} [key] - Chave opcional.
  */

  /**
   * @returns {void}
   */
  cacheAllData() {
    this._lastRefreshInCache = new Date();
    const {headerRow, lastColumn, lastRow, sheet} = this;
    this._data = sheet.getRange(1,1,lastRow,lastColumn).getValues();

    for(let rowNumber = headerRow + 1; rowNumber <= lastRow; rowNumber++){
      this.updateMapByPrimaryKey(rowNumber);
    }
  }

  /**
   * @param {{rowNumber:number, rowData:any[]}} param0
   * @returns {void}
   */
  cacheRowData({rowNumber, rowData}){
    this.data[rowNumber-1] = rowData;
    this.updateMapByPrimaryKey(rowNumber);
  }

  /**
   * @param {number} rowNumber
   * @returns {void}
   */
  updateMapByPrimaryKey(rowNumber) {
    const rowData = this.data[rowNumber-1];

    const key = this.primaryKey
      .reduce((key, colName, index) => {
        const colNumber = this.getColNumber(colName);
        const section = colNumber ? rowData[colNumber-1] : colName === 'rowNumber' ? rowNumber : '';
        return section ? key + (index !== 0 ? '&' : '') + section : key;
      }, '');

    this.mapByPrimaryKey[key] = rowNumber;
  }

  clearCache() {
    this._data = [];
    this._lastColumn = undefined;
    this._lastRow = undefined;
    this._mapByPrimaryKey = {};
  }

  /**
   * @param {number} colNumber
   * @returns {string}
   */
  getColNameByColNumber(colNumber) {
    const { className, headerRowData } = this;

    if (colNumber <= 0) {
      throw (
        `Ops! Erro ao chamar "${className}.getColNameByColNumber()":\n` +
        'O parâmetro "colNumber" deve ser um número maior que zero.'
      );
    }
    return headerRowData[colNumber - 1];
  }

  /**
   * @param {string} colName
   * @returns {number}
   */
  getColNumber(colName) {
    const { columnMap, headerRowData } = this;
    colName = (columnMap && columnMap[colName]) || colName;
    const index = headerRowData.indexOf(colName);
    return index >= 0 ? index + 1 : undefined;
  }

  /**
   * @param {string} colName
   * @returns {number}
   */
  getMaxInColumn(colName) {
    const { headerRow, sheet, lastRow } = this;
    const initialRow = headerRow + 1;
    const column = this.getColNumber(colName);

    return sheet
      .getRange(initialRow, column, lastRow)
      .getValues()
      .sort((a, b) => b[0] - a[0])[0][0];
  }

  /**
   * Cria um arquivo PDF com o conteúdo da planilha.
   * @param {string} fileName O nome do arquivo que será criado
   * @returns {File} O arquivo PDF criado
   */
  getPDFFromSheet(fileName) {
    const { sheet, ssId } = this;
    return SSheet.getPDFFromSheet(fileName, sheet, ssId);
  }

  /**
   * @param {object} obj
   * @returns {string}
   */
  getPrimaryMapKey(obj) {
    const result = this.primaryKey
      .map((colName) => obj[colName] || '')
      .join('&');
    return result;
  }

  /**
   * @param {string} rangeA1
   * @returns {Range}
   */
  getRange(rangeA1) {
    return this.sheet.getRange(rangeA1);
  }

  /**
   * @param {string} rangeA1
   * @returns {any}
   */
  getRangeVal(rangeA1) {
    return this.getRange(rangeA1).getValue();
  }

  /**
   * @param {number} rowNumber
   * @returns {any[]}
   */
  getRowData(rowNumber) {
    const { sheet, lastColumn, lastRow } = this;

    if (
      !(Number.isInteger(rowNumber) && rowNumber > 0 && rowNumber <= lastRow)
    ) {
      throw (
        `Ops! Erro ao chamar "${this.className}.getRowData()":\n` +
        'O parâmetro "rowNumber" deve ser um número maior que zero e menor ou igual à última linha da planilha.'
      );
    }

    return sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  }

  /**
   * @param {RowObject} rowObject
   */
  getRowDataByRowObject(rowObject) {
    return this.headerRowData.map((colName) => {
      let value = rowObject[colName];
      return value !== undefined ? value : '';
    });
  }

  /**
   * @param {{rowNumber:number, rowData:any[]}} object
   * @returns {RowObject}
   */
  getRowObjectByRowData({rowNumber, rowData}) {
    if (Array.isArray(rowData)) {
      return this.headerRowData.reduce(
        (obj, colName, index) => {
          const value = rowData[index];
          obj[colName] = value === '' ? undefined : value;

          return obj;
        },
        { rowNumber }
      );
    }
    throw 'Ops! O parâmetro "rowData" deve ser um array.';
  }

  /**
   * @param {Range} range
   * @returns {any[][]}
   */
  getFormulas(range) {
    return mergeThreeDimensionalArrays(
      range.getFormulas(),
      range.getFormulasR1C1()
    );
  }

  /**
   * @param {object} e
   * @param {function} callback
   * @returns {void}
   */
  handleEditEvent(e, callback = null) {
    const { oldValue, range, source: spreadsheet, value } = e;

    const sheetName = range.getSheet().getName();
    const ssId = spreadsheet.getId();
    const sheet = range.getSheet();

    if (ssId !== this.ssId || sheetName !== this.sheetName) return;

    const rowNumber = range.getRow();
    const colNumber = range.getColumn();

    let colName;

    if (value !== undefined || oldValue !== undefined) {
      const { headerRow } = this;

      if (headerRow && rowNumber > headerRow) {
        colName = this.getColNameByColNumber(colNumber);
      }
    }
    try {
      if (callback) {
        callback({
          colName,
          colNumber,
          oldValue,
          rowNumber,
          range,
          sheet,
          sheetName,
          spreadsheet,
          ssId,
          value,
        });
      }
    } catch(e) {
      //
    }
  }

  /**
   * @returns {void}
   */
  hide() {
    this.sheet.hideSheet();
  }

  /**
  * @returns{boolean}
  */
  isCacheCompleted() {
    const {data, headerRow, lastRow} = this;
    return data[headerRow-1] && data.length === lastRow;
  }

  /**
   * @returns {boolean}
   */
  isInTransaction() {
    return !!this._inTransaction;
  }

  /**
   * @returns {boolean}
   */
  isRowNumberPrimaryKey() {
    const { primaryKey } = this;
    return primaryKey.length == 1 && primaryKey[0] == 'rowNumber';
  }

  /**
   * @param {object} [query]
   * @param {number} [rowNumber]
   * @param {number} [limit]
   * @returns {Model[]|RowObject[]}
   */
  read(query = {}, rowNumber = null, limit = null) {
    const {
      className,
      columnMap,
      entity,
      headerRow,
      mapByPrimaryKey,
      primaryKey,
    } = this;

    /** @type {Model[] | RowObject[] | object[]} result */
    let result = [];

    //ATTENTION! When rowNumber is passed with query,
    //the search starts from that point, improving performance.
    if (rowNumber && !query) {
      if (rowNumber <= headerRow) {
        throw (
          `Ops! Erro ao chamar ${className}.read():\n` +
          'O parâmetro "rowNumber" deve ser maior que o número da linha do cabeçalho.'
        );
      }

      let rowData = this.data[rowNumber-1];

      if(!rowData){
        rowData = this.getRowData(rowNumber);
        this.cacheRowData({rowNumber, rowData});
      }

      const rowObject = this.getRowObjectByRowData({rowNumber, rowData});

      if (!rowObject) {
        throw (
          `Ops! Erro ao chamar ${className}.read():\n` +
          `Não foi encontrado nenhum dado na linha ${rowNumber}`
        );
      }
      result.push(rowObject);
    } else {
      let found = 0;
      const queryKeys = [];

      if (columnMap) {
        query = Object.entries(query).reduce((obj, [key, value]) => {
          const newKey = columnMap[key] || key;
          obj[newKey] = value;
          queryKeys.push(newKey);
          return obj;
        }, {});
      }

      const queryHasAllPrimaryKeys = primaryKey.every(
        colName => query[colName] !== undefined && typeof query[colName] !== 'function'
      );

      /**
       * @param {number} rowNumber
       */
      const pushToResultIfMatches = (rowNumber) => {
        const rowData = this.data[rowNumber-1];

        const matched = Object.entries(query).reduce((matched, [colName, searchValue]) => {
          if(matched){
            const colNumber = this.getColNumber(colName);
            let foundValue = rowData[colNumber-1];

            if(foundValue === '') foundValue = undefined;

            return isDate(foundValue) && isDate(searchValue)
              ? isSameMinute(foundValue, searchValue)
              : typeof searchValue === 'function'
                ? searchValue(foundValue)
                : foundValue === searchValue;
          }
          return matched;
        }, true);

        if(matched){
          result.push(this.getRowObjectByRowData({rowNumber, rowData}));
          found++;
        }
      }

      if (queryHasAllPrimaryKeys) {
        const primaryMapKey = this.getPrimaryMapKey(query);

        rowNumber = primaryMapKey && mapByPrimaryKey[primaryMapKey];

        if(rowNumber){
          pushToResultIfMatches(rowNumber);
        }else if(!this.isCacheCompleted()){
          this.cacheAllData();
          return this.read(query, null, 1);
        }
      } else {
        if (!this.isCacheCompleted()) {
          this.cacheAllData();
        }
        limit = limit || 10000;

        let len = this.data.length;

        for (rowNumber = headerRow; rowNumber <= len && found < limit; rowNumber++) {
          pushToResultIfMatches(rowNumber);
        }
      }
    }

    if (columnMap) {
      result = result.map((rowObject) => {
        const { rowNumber } = rowObject;
        /**@type {object|Model} */
        let model = Object.entries(columnMap).reduce(
          (obj, [key, colName]) => {
            obj[key] = rowObject[colName];
            return obj;
          },
          { rowNumber }
        );
        if (entity) {
          model = Model.getFromJSON(entity, model);
          model.rowNumber = rowNumber;
          model = model.makeObservable();
        }
        return model;
      });
    }
    return result;
  }

  /**
   * @returns {SSheet}
   */
  refreshFilterViews() {
    const { lastColumn, lastRow, sheetId, sheetName, ssId } = this;

    try {
      const filterViews = Sheets.Spreadsheets.get(ssId, {
        ranges: [sheetName],
        fields: 'sheets(filterViews)',
      }).sheets[0].filterViews;

      if (filterViews && filterViews.length) {
        const requests = filterViews.map((e) => ({
          updateFilterView: {
            filter: {
              filterViewId: e.filterViewId,
              range: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: lastRow,
                startColumnIndex: 0,
                endColumnIndex: lastColumn,
              },
            },
            fields: '*',
          },
        }));
        Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
      }
    } catch (e) {
      console.error(e);
    }
    return this;
  }

  /**
   * @param {object} object
   */
  render(object = null) {
    if (!this.rangesA1) {
      throw (
        `Ops! Erro ao chamar '${this.className}.render()'. ` +
        "O parâmetro 'rangesA1' não foi definido."
      );
    }
    Object.entries(this.rangesA1).forEach(([key, rangeA1]) => {
      try {
        const value = object && object[key];
        if (value !== undefined) {
          this.sheet.getRange(rangeA1).setValue(value);
        }
      } catch (e) {
        //Tudo bem!
      }
    });
  }

  /**
   * @param {Model|Model[]|object|object[]} data
   * @returns {void}
   */
  save(data) {
    const { columnMap, entity, lastRefreshInCache, primaryKey, readOnlyKeys, sheet } = this;

    /** @type {any[][]} */
    const valuesToAppend = [];

    /** @type {Object<string,any[]>} */
    const mapOfRowsToUpdate = {};

    /** @type {number} */
    let minCol = undefined;

    /** @type {number} */
    let maxCol = undefined;

    data = Array.isArray(data) ? data : [data];

    /**
     * @param {Model} model
     * @returns {RowObject}
     */
    const getRowObjectByModel = (model) => {
      if (!(model instanceof entity)) {
        throw (
          `Ops! Erro ao chamar "${this.className}.save()". ` +
          `O parâmetro "data" deve ser um array de objetos do tipo "${entity.name}".`
        );
      }

      model.init();
      model.validate();

      const { rowNumber } = model;

      const modelIsObservable = Model.isObservable(model);
      const getFullRowObject = !(rowNumber && modelIsObservable);

      if(columnMap){
        return Object.entries(columnMap).reduce(
          (obj, [key, columnName]) => {
            if (
              getFullRowObject ||
              primaryKey.includes(columnName) ||
              (modelIsObservable && model.hasChanged(key))
            ) {
              obj[columnName] = model[key];
            }
            return obj;
          },
          { rowNumber }
        );
      }else{
        return model.toJSON();
      }
    };

    /**
     * @param {RowObject} rowObject
     */
    const updateCachedRowData = (rowObject) => {
      let { rowNumber } = rowObject;

      let quantityOfQueryKeys = 0;

      let query = rowNumber ? null
        : primaryKey.reduce((obj, colName) => {
            const newValue = rowObject[colName];
            if(newValue !== undefined){
              quantityOfQueryKeys++;
              obj[colName] = rowObject[colName];
            }
            return obj;
          }, {});

      if(!(rowNumber || quantityOfQueryKeys)){
        throw (
          `Ops! Erro ao chamar "${this.className}.save()":\n` +
          'Não foi possível encontrar a linha a ser atualizada. ' +
          'Você deve informar o número da linha ou o valor de pelo menos ' +
          'uma das colunas que compõem a chave primária.'
        );
      }

      const oldRowObject = this.read(query, rowNumber, 1)[0];

      let rowData;

      if(!oldRowObject){
        rowData = this.getRowDataByRowObject(rowObject);
        valuesToAppend.push(rowData);
        return;
      }

      ({rowNumber} = oldRowObject);

      let dataChanged = false;

      rowData = this.data[rowNumber-1];

      Object.entries(rowObject).forEach(([colName, newValue]) => {
        const colNumber = this.getColNumber(colName);
        const oldValue = colNumber ? rowData[colNumber-1] : undefined;

        if([undefined, null].includes(newValue))
          newValue = '';

        if (colNumber && (!readOnlyKeys.includes(colName) || oldValue === undefined)) {
          const valueHasChanged =
            newValue instanceof Date && oldValue instanceof Date
              ? !isSameMinute(newValue, oldValue)
              : newValue !== oldValue;

          if (valueHasChanged) {
            if (primaryKey.includes(colName)) {
              throw (
                `Ops! Erro ao chamar "${this.className}.save()":\n` +
                `Você não pode alterar o valor da coluna "${colName}" de ` +
                `"${oldValue}" para "${newValue}" porque ela faz parte da chave primária.`
              );
            }
            if(!minCol || minCol > colNumber)
              minCol = colNumber;

            if(!maxCol || maxCol < colNumber)
              maxCol = colNumber;

            rowData[colNumber-1] = newValue;
            dataChanged = true;
          }
        }
      });

      if(dataChanged){
        mapOfRowsToUpdate[rowNumber] = rowData;
      }
    }

    //IMPORTANTE! Se o parâmetro "data" for um array de objetos do tipo Model,
    //além de iniciá-los e validá-los, é necessário converter para um array de
    //objetos do tipo RowObject,
    if (entity) {
      data = data.map((/** @type {Model} */ model) => getRowObjectByModel(model));
    }

    //Se a quantidade de linhas a serem atualizadas for maior que 20,
    //é mais eficiente atualizar todas as linhas de uma só vez. Do
    //contrário, atualizaremos os dados linha por linha.
    const mustSetAllValuesAtOnce = data.length > 20;

    //IMPORTANTE! Pega valores atualizados para evitar que
    //dados sejam gravados em linhas ou colunas erradas caso
    //tenham sido inseridas ou excluídas linhas ou colunas
    if(!lastRefreshInCache || differenceInSeconds(new Date(), lastRefreshInCache) >= 60){
      this.clearCache();
    }

    if(mustSetAllValuesAtOnce && !this.isCacheCompleted()){
      this.cacheAllData();
    }

    //É necessário limpar o columnMap temporariamente para forçar
    //que o resultado de SSheet.read() seja um rowObject, e não um model
    this.columnMap = null;

    data.forEach((/** @type {RowObject} */ newRowObject) => {
      updateCachedRowData(newRowObject);
    });

    if (!this.isInTransaction()) {
      /**
       * @param {number} row
       * @param {number} column
       * @param {any[][]} values
       */
      const safelySetValues = (row, column, values) => {
        values = values.map((rowData) => {
          return rowData.map((value, index) => {
            const colName = this.getColNameByColNumber(column + index);
            return this.columnsWithMapFormula.includes(colName) ? '' : value;
          });
        })
        const range = sheet.getRange(row, column, values.length, values[0].length);
        range.setValues(values);
      }

      if (valuesToAppend.length) {
        this._lastRow = null;
        safelySetValues(this.lastRow + 1, 1, valuesToAppend);
      }

      const rowsToUpdate = Object.keys(mapOfRowsToUpdate).map(key => parseInt(key));

      if(!mustSetAllValuesAtOnce){
        let lastRow;
        let currentArr = [];
        let size = rowsToUpdate.length;

        rowsToUpdate.reduce((arr, currentRow, index) => {
          const isLastItem = index == size - 1;

          if(!lastRow || lastRow == currentRow -1){
            currentArr.push(currentRow);
          }else{
            arr.push(currentArr);
            currentArr = [currentRow];
          }
          if(isLastItem){
            arr.push(currentArr);
          }
          lastRow = currentRow;
          return arr;
        }, []).forEach((arr) => {
          const row = arr[0];
          const values = arr.map((/**@type {number} */ rowNumber) => mapOfRowsToUpdate[rowNumber].slice(minCol-1, maxCol));
          safelySetValues(row, minCol, values);
        });
      }else{
        const minRow = Math.min(...rowsToUpdate);
        const maxRow = Math.max(...rowsToUpdate);

        const values = this.data.slice(minRow-1, maxRow)
          .map((rowData) => rowData.slice(minCol-1, maxCol));

        safelySetValues(minRow, minCol, values);
      }
    }else{
      //necessário implementação
    }
    //Redefine o columnMap
    this.columnMap = columnMap;
    this.clearCache();
  }

  /**
   * @param {string} key
   * @param {any} value
   * @returns {SSheet}
   */
  set(key, value) {
    const rangeA1 = this.rangesA1[key];

    if (rangeA1) this.setRangeVal(rangeA1, value);

    return this;
  }

  /**
   * @param {string} rangeA1
   * @param {any} value
   * @returns {SSheet}
   */
  setRangeVal(rangeA1, value) {
    this.getRange(rangeA1).setValue(value);
    return this;
  }

  /**
   * @returns {void}
   */
  show() {
    this.sheet.showSheet().activate();
  }

  /**
   * @param {string} ssId
   * @param {number} tries
   * @returns {Spreadsheet}
   */
  static safelyOpenSpreadsheetById(ssId, tries = 0) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      if (++tries <= 10) {
        Utilities.sleep(Math.pow(2, tries - 1) * 100);
        return SSheet.safelyOpenSpreadsheetById(ssId, tries);
      } else {
        const msg = `Não foi possível abrir a Spreadsheet com ID '${ssId}' após 10 tentativas`;
        console.error(`${msg}. Erro:\n${e}`);
        throw `Ops! ${msg}`;
      }
    }
  }

  /**
   *
   * @param {Spreadsheet} ss
   * @param {string} sheetName
   * @param {number} tries
   * @returns {Sheet}
   */
  static safelyGetSheetByName(ss, sheetName, tries = 0) {
    tries++;
    try {
      return ss.getSheetByName(sheetName);
    } catch (e) {
      if (tries <= 10) {
        Utilities.sleep(Math.pow(2, tries - 1) * 100);

        return SSheet.safelyGetSheetByName(ss, sheetName, tries);
      } else {
        const msg = `Não foi possível abrir a planilha com nome '${sheetName}' após 10 tentativas`;
        console.error(`${msg}. Erro:\n${e}`);
        throw `Ops! ${msg}`;
      }
    }
  }

  /**
   * Cria um arquivo PDF com o conteúdo de uma planilha.
   * @param {string} fileName O nome do arquivo que será criado
   * @returns {File} O arquivo PDF criado
   */
  static getPDFFromSheet(fileName, sheet, ssId) {
    SpreadsheetApp.flush();

    const sheetIsHidden = sheet.isSheetHidden();

    //Se a planilha estava oculta, é necessário mostrá-la para gerar o PDF
    sheet.showSheet().activate();

    const ss = sheet.getParent();

    const parents = DriveApp.getFileById(ssId).getParents();
    const folder = parents.hasNext()
      ? parents.next()
      : DriveApp.getRootFolder();

    const url = ss.getUrl();

    const exportUrl =
      url.replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +
      '&top_margin=0.75' +
      '&bottom_margin=0.75' +
      '&left_margin=0.7' +
      '&right_margin=0.7' +
      '&sheetnames=false&printtitle=false' +
      '&pagenum=false' +
      '&gridlines=true' +
      '&fzr=FALSE' +
      '&gid=' +
      sheet.getSheetId();

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
    });
    const blob = response.getBlob();

    if (sheetIsHidden) sheet.hideSheet();

    return folder.createFile(blob.setName(fileName));
  }
}
