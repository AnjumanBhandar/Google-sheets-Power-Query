// ============================================================
//  Google Sheets Power Query  –  Engine.gs
//  Pure transform pipeline. No SpreadsheetApp calls here.
//  Input/output: 2-D arrays (rows × cols), all values as-is.
// ============================================================

var Engine = (function() {

  // ── Public entry point ─────────────────────────────────────

  /**
   * Run the full pipeline.
   * @param {Array[][]} raw       - Full raw sheet data (including headers)
   * @param {number}    headerRows - 1 or 2 header rows
   * @param {Object[]}  steps     - Array of step definitions
   * @returns {Array[][]} Transformed data with a single header row
   */
  function run(raw, headerRows, steps) {
    if (!raw || raw.length === 0) return [];

    headerRows = headerRows || 1;

    // Build a single merged header row
    var header = _buildHeader(raw, headerRows);
    var data   = raw.slice(headerRows);

    // Remove completely empty rows from source
    data = data.filter(function(row) {
      return row.some(function(cell) { return cell !== '' && cell !== null && cell !== undefined; });
    });

    // Execute each step in sequence
    for (var i = 0; i < steps.length; i++) {
      var step = steps[i];
      if (!step || step.disabled) continue;

      switch (step.type) {
        case 'unpivot':         var r = _unpivot(header, data, step.params);        header = r.header; data = r.data; break;
        case 'filter':          data  = _filter(header, data, step.params);         break;
        case 'selectCols':      var r = _selectCols(header, data, step.params);     header = r.header; data = r.data; break;
        case 'sort':            data  = _sort(header, data, step.params);           break;
        case 'addColumn':       var r = _addColumn(header, data, step.params);      header = r.header; data = r.data; break;
        case 'changeType':      data  = _changeType(header, data, step.params);     break;
        case 'removeDuplicates':data  = _removeDuplicates(header, data, step.params); break;
        case 'trimClean':       data  = _trimClean(header, data, step.params);      break;
        default: break; // unknown step – skip silently
      }
    }

    // Prepend header and return
    return [header].concat(data);
  }


  // ── Header builder ─────────────────────────────────────────

  function _buildHeader(raw, headerRows) {
    if (headerRows === 1) {
      return raw[0].map(function(v) { return String(v).trim() || '(blank)'; });
    }
    // 2-row merge: "Q1" + "Sales" → "Q1 Sales"
    var r1 = raw[0], r2 = raw[1];
    var merged = [];
    var lastGroup = '';
    for (var c = 0; c < Math.max(r1.length, r2.length); c++) {
      var top = String(r1[c] || '').trim();
      var bot = String(r2[c] || '').trim();
      if (top) lastGroup = top;
      if (top && bot)  merged.push(top + ' ' + bot);
      else if (top)    merged.push(top);
      else if (bot)    merged.push(lastGroup ? lastGroup + ' ' + bot : bot);
      else             merged.push('Col' + (c + 1));
    }
    return merged;
  }


  // ── Step: Unpivot ──────────────────────────────────────────
  /*
    params: {
      anchorCols:  ['Region', 'Product'],   // columns that stay fixed
      pivotCols:   ['Jan', 'Feb', 'Mar'],   // columns to unpivot
      labelColName: 'Month',
      valueColName: 'Value'
    }
  */
  function _unpivot(header, data, params) {
    var anchors  = params.anchorCols  || [];
    var pivots   = params.pivotCols   || [];
    var labelCol = params.labelColName || 'Attribute';
    var valueCol = params.valueColName || 'Value';

    // If no pivot cols specified, pivot everything that isn't an anchor
    if (pivots.length === 0) {
      pivots = header.filter(function(h) { return anchors.indexOf(h) === -1; });
    }

    var anchorIdx = anchors.map(function(a) { return header.indexOf(a); });
    var pivotIdx  = pivots.map(function(p)  { return header.indexOf(p); });

    var newHeader = anchors.concat([labelCol, valueCol]);
    var newData   = [];

    data.forEach(function(row) {
      pivotIdx.forEach(function(pi, i) {
        if (pi === -1) return;
        var newRow = anchorIdx.map(function(ai) { return ai >= 0 ? row[ai] : ''; });
        newRow.push(pivots[i]);
        newRow.push(pi < row.length ? row[pi] : '');
        newData.push(newRow);
      });
    });

    return { header: newHeader, data: newData };
  }


  // ── Step: Filter ───────────────────────────────────────────
  /*
    params: {
      rules: [
        { col: 'Region', op: 'equals',    value: 'North' },
        { col: 'Sales',  op: 'gt',        value: 1000 },
        { col: 'Notes',  op: 'notBlank',  value: '' }
      ],
      logic: 'AND'  // 'AND' | 'OR'
    }
  */
  function _filter(header, data, params) {
    var rules = params.rules || [];
    var logic = (params.logic || 'AND').toUpperCase();
    if (rules.length === 0) return data;

    return data.filter(function(row) {
      var results = rules.map(function(rule) {
        var ci  = header.indexOf(rule.col);
        if (ci === -1) return true;
        var cell = row[ci];
        var val  = rule.value;
        switch (rule.op) {
          case 'equals':      return String(cell).toLowerCase() === String(val).toLowerCase();
          case 'notEquals':   return String(cell).toLowerCase() !== String(val).toLowerCase();
          case 'contains':    return String(cell).toLowerCase().indexOf(String(val).toLowerCase()) !== -1;
          case 'notContains': return String(cell).toLowerCase().indexOf(String(val).toLowerCase()) === -1;
          case 'startsWith':  return String(cell).toLowerCase().indexOf(String(val).toLowerCase()) === 0;
          case 'gt':          return Number(cell) > Number(val);
          case 'gte':         return Number(cell) >= Number(val);
          case 'lt':          return Number(cell) < Number(val);
          case 'lte':         return Number(cell) <= Number(val);
          case 'blank':       return cell === '' || cell === null || cell === undefined;
          case 'notBlank':    return cell !== '' && cell !== null && cell !== undefined;
          default:            return true;
        }
      });
      return logic === 'OR'
        ? results.some(Boolean)
        : results.every(Boolean);
    });
  }


  // ── Step: Select / reorder / rename columns ────────────────
  /*
    params: {
      cols: [
        { original: 'Region', rename: 'Territory' },
        { original: 'Sales',  rename: '' },
        { original: 'Month',  rename: '' }
      ]
    }
  */
  function _selectCols(header, data, params) {
    var cols = params.cols || [];
    if (cols.length === 0) return { header: header, data: data };

    var indices = cols.map(function(c) { return header.indexOf(c.original); });
    var newHeader = cols.map(function(c, i) {
      return (c.rename && c.rename.trim()) ? c.rename.trim() : c.original;
    });

    var newData = data.map(function(row) {
      return indices.map(function(idx) { return idx >= 0 ? row[idx] : ''; });
    });

    return { header: newHeader, data: newData };
  }


  // ── Step: Sort ─────────────────────────────────────────────
  /*
    params: {
      keys: [
        { col: 'Region', dir: 'asc' },
        { col: 'Sales',  dir: 'desc' }
      ]
    }
  */
  function _sort(header, data, params) {
    var keys = params.keys || [];
    if (keys.length === 0) return data;

    var keyDefs = keys.map(function(k) {
      return { idx: header.indexOf(k.col), dir: k.dir === 'desc' ? -1 : 1 };
    });

    return data.slice().sort(function(a, b) {
      for (var i = 0; i < keyDefs.length; i++) {
        var kd = keyDefs[i];
        if (kd.idx === -1) continue;
        var av = a[kd.idx], bv = b[kd.idx];
        // Numeric sort if both look like numbers
        var an = Number(av), bn = Number(bv);
        if (!isNaN(an) && !isNaN(bn)) {
          if (an !== bn) return kd.dir * (an - bn);
        } else {
          var as_ = String(av).toLowerCase(), bs_ = String(bv).toLowerCase();
          if (as_ < bs_) return kd.dir * -1;
          if (as_ > bs_) return kd.dir *  1;
        }
      }
      return 0;
    });
  }


  // ── Step: Add custom column ────────────────────────────────
  /*
    params: {
      newColName: 'Full Name',
      expression: 'concat',      // 'concat' | 'math' | 'text' | 'conditional'
      // For concat:
      cols: ['First', 'Last'],
      separator: ' ',
      // For math:
      col1: 'Revenue', op: '+', col2: 'Bonus',   // or staticVal: 100
      // For text:
      col: 'Name', textOp: 'upper',              // 'upper'|'lower'|'trim'|'len'
      // For conditional:
      ifCol: 'Score', ifOp: 'gte', ifVal: '50', thenVal: 'Pass', elseVal: 'Fail'
    }
  */
  function _addColumn(header, data, params) {
    var newName = params.newColName || 'New Column';
    var newHeader = header.concat([newName]);

    var newData = data.map(function(row) {
      var val = _evalExpression(header, row, params);
      return row.concat([val]);
    });

    return { header: newHeader, data: newData };
  }

  function _evalExpression(header, row, params) {
    var get = function(colName) {
      var i = header.indexOf(colName);
      return i >= 0 ? row[i] : '';
    };

    switch (params.expression) {
      case 'concat':
        var parts = (params.cols || []).map(get);
        return parts.join(params.separator !== undefined ? params.separator : ' ');

      case 'math':
        var v1 = Number(get(params.col1));
        var v2 = params.col2 ? Number(get(params.col2)) : Number(params.staticVal || 0);
        switch (params.op) {
          case '+': return v1 + v2;
          case '-': return v1 - v2;
          case '*': return v1 * v2;
          case '/': return v2 !== 0 ? v1 / v2 : '';
          case '%': return v2 !== 0 ? v1 % v2 : '';
          default:  return '';
        }

      case 'text':
        var tv = String(get(params.col));
        switch (params.textOp) {
          case 'upper': return tv.toUpperCase();
          case 'lower': return tv.toLowerCase();
          case 'trim':  return tv.trim();
          case 'len':   return tv.length;
          default:      return tv;
        }

      case 'conditional':
        var cv   = get(params.ifCol);
        var cval = params.ifVal || '';
        var pass = false;
        switch (params.ifOp) {
          case 'equals':   pass = String(cv).toLowerCase() === String(cval).toLowerCase(); break;
          case 'contains': pass = String(cv).toLowerCase().indexOf(String(cval).toLowerCase()) !== -1; break;
          case 'gt':       pass = Number(cv) > Number(cval);  break;
          case 'gte':      pass = Number(cv) >= Number(cval); break;
          case 'lt':       pass = Number(cv) < Number(cval);  break;
          case 'lte':      pass = Number(cv) <= Number(cval); break;
          case 'blank':    pass = cv === '' || cv === null || cv === undefined; break;
          case 'notBlank': pass = cv !== '' && cv !== null && cv !== undefined; break;
        }
        return pass ? params.thenVal : params.elseVal;

      default:
        return '';
    }
  }


  // ── Step: Change data type ─────────────────────────────────
  /*
    params: {
      conversions: [
        { col: 'Sales',      toType: 'number' },
        { col: 'Date',       toType: 'date',   dateFormat: 'yyyy-MM-dd' },
        { col: 'Active',     toType: 'boolean' },
        { col: 'Notes',      toType: 'text' }
      ]
    }
  */
  function _changeType(header, data, params) {
    var convs = params.conversions || [];

    return data.map(function(row) {
      var newRow = row.slice();
      convs.forEach(function(conv) {
        var ci = header.indexOf(conv.col);
        if (ci === -1) return;
        var v = newRow[ci];
        switch (conv.toType) {
          case 'number':
            newRow[ci] = Number(String(v).replace(/[^0-9.\-]/g, '')) || 0;
            break;
          case 'text':
            newRow[ci] = String(v);
            break;
          case 'boolean':
            var s = String(v).toLowerCase().trim();
            newRow[ci] = (s === 'true' || s === 'yes' || s === '1') ? true : false;
            break;
          case 'date':
            var d = new Date(v);
            newRow[ci] = isNaN(d.getTime()) ? v : _formatDate(d, conv.dateFormat || 'yyyy-MM-dd');
            break;
        }
      });
      return newRow;
    });
  }

  function _formatDate(d, fmt) {
    var y  = d.getFullYear();
    var m  = String(d.getMonth() + 1).padStart(2, '0');
    var dd = String(d.getDate()).padStart(2, '0');
    return fmt
      .replace('yyyy', y).replace('yy', String(y).slice(-2))
      .replace('MM', m).replace('dd', dd);
  }


  // ── Step: Remove duplicates ────────────────────────────────
  /*
    params: {
      keyCols: ['Region', 'Month']  // empty = use all cols
    }
  */
  function _removeDuplicates(header, data, params) {
    var keyCols = params.keyCols || [];
    var keyIdx  = keyCols.length > 0
      ? keyCols.map(function(c) { return header.indexOf(c); })
      : header.map(function(_, i) { return i; });

    var seen = {};
    return data.filter(function(row) {
      var key = keyIdx.map(function(i) { return i >= 0 ? String(row[i]) : ''; }).join('|||');
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
  }


  // ── Step: Trim / clean text ────────────────────────────────
  /*
    params: {
      cols: ['Name', 'Notes'],   // empty = all cols
      ops: ['trim', 'lower', 'upper', 'removeExtraSpaces', 'removeSpecial']
    }
  */
  function _trimClean(header, data, params) {
    var targetCols = params.cols || [];
    var ops        = params.ops  || ['trim'];

    var targetIdx = targetCols.length > 0
      ? targetCols.map(function(c) { return header.indexOf(c); })
      : header.map(function(_, i) { return i; });

    return data.map(function(row) {
      var newRow = row.slice();
      targetIdx.forEach(function(ci) {
        if (ci === -1 || typeof newRow[ci] !== 'string') return;
        var v = newRow[ci];
        ops.forEach(function(op) {
          switch (op) {
            case 'trim':             v = v.trim(); break;
            case 'lower':            v = v.toLowerCase(); break;
            case 'upper':            v = v.toUpperCase(); break;
            case 'removeExtraSpaces':v = v.replace(/\s+/g, ' ').trim(); break;
            case 'removeSpecial':    v = v.replace(/[^a-zA-Z0-9 ]/g, ''); break;
          }
        });
        newRow[ci] = v;
      });
      return newRow;
    });
  }


  // ── Public API ─────────────────────────────────────────────
  return { run: run };

})();