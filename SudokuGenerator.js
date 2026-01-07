Array.prototype.pick = function() {
    if (!this.length)
        return null;
    
    var r = Math.round(Math.random() * (this.length - 1));
    return this.splice(r, 1)[0];
};

Math.randomEx = function(from, to) {
    return Math.round(Math.random() * (to - from)) + from;
};

Array.prototype.clear = function() {
    this.splice(0, this.length);
};

var PER_SUDOKU = 0;
var PER_GRID = 1;
var PER_BOX = 2;
var PER_ROW = 3;
var PER_COLUMN = 4;

var $CENTRAL = 2;
var $UPPER_LEFT = 0;
var $UPPER_RIGHT = 1;
var $LOWER_RIGHT = 4;
var $LOWER_LEFT = 3;

with (WSH) {
    var namedArgs = Arguments.Named;
    if (namedArgs.Exists("NS")) {
        var n = parseInt(namedArgs("NS"));
        if (isNaN(n) || (n < 0)) {
            beep(StdErr);
            StdErr.WriteLine("[error] The /NS parameter not set to a valid unsigned integer.");
            Quit(1);
        }
    }
    else
        var n = 1;
    
    if (namedArgs.Exists("NG")) {
        var $n = parseInt(namedArgs("NG"));
        if (isNaN($n) || ($n < 0)) {
            beep(StdErr);
            StdErr.WriteLine("[error] The /NG parameter not set to a valid unsigned integer.");
            Quit(1);
        }
    }
    else
        var $n = 0;
    
    if (namedArgs.Exists("EP")) {
        var ep = namedArgs("EP");
        var portion = {};
        var maxPortionSize = 0;
        if (!ep) {
            beep(StdErr);
            StdErr.WriteLine("[error] The /EP parameter contains no value.");
            Quit(1);
        }
        
        switch (ep.slice(-1)) {
        case "B" :
            portion.scope = PER_BOX;
            maxPortionSize = 9;
            break;
        case "R" :
            portion.scope = PER_ROW;
            maxPortionSize = 9;
            break;
        case "C" :
            portion.scope = PER_COLUMN;
            maxPortionSize = 9;
            break;
        case "G" :
            portion.scope = PER_GRID;
            maxPortionSize = $n ? 72 : 81;
            break;
        default :
            portion.scope = PER_SUDOKU;
            maxPortionSize = n ? 81 : 369;
        }
        
        var tokens = ep.split("-");
        if (tokens.length == 1) {
            var s = parseInt(tokens[0]);
            if (isNaN(s) || (s < 0) || (s > maxPortionSize)) {
                beep(StdErr);
                StdErr.WriteLine("[error] Invalid usage of /EP parameter.");
                Quit(1);
            }
            
            portion.size = s;
        }
        else if (tokens.length == 2) {
            var lb = parseInt(tokens[0]);
            if (isNaN(lb) || (lb < 0) || (lb > maxPortionSize)) {
                beep(StdErr);
                StdErr.WriteLine("[error] Invalid usage of /EP parameter.");
                Quit(1);
            }
            
            var ub = parseInt(tokens[1]);
            if (isNaN(ub) || (ub < 0) || (ub > maxPortionSize)) {
                beep(StdErr);
                StdErr.WriteLine("[error] Invalid usage of /EP parameter.");
                Quit(1);
            }
            
            if (lb > ub) {
                beep(StdErr);
                StdErr.WriteLine("[error] Invalid usage of /EP parameter.");
                Quit(1);
            }
            
            portion.lBound = lb;
            portion.uBound = ub;
        }
        else {
            beep(StdErr);
            StdErr.WriteLine("[error] Invalid usage of /EP parameter.");
            Quit(1);
        }
    }
    else
        var portion = { size: 0, scope: PER_SUDOKU };
    
    var nFailed = 0;
    for (var i = 0; i < n; i++) {
        try {
            var sudoku = generateSudoku(portion);
            
            if (i)
                StdOut.WriteBlankLines(1);
            
            outputSudoku(StdOut, sudoku);
        }
        catch (err) {
            StdErr.WriteLine("[warn] " + err.message);
            nFailed++;
        }
    }
    
    var $nFailed = 0;
    for (var i = 0; i < $n; i++) {
        try {
            var $sudoku = $generateSudoku(portion);
            
            StdOut.WriteBlankLines(1);
            $outputSudoku(StdOut, $sudoku);
        }
        catch (err) {
            StdErr.WriteLine("[warn] " + err.message);
            $nFailed++;
        }
    }
    
    if (nFailed) {
        beep(StdErr);
        StdErr.WriteLine("[warn] Failed to generate " + nFailed + " of the normal sudokus requested.");
    }
    
    if ($nFailed) {
        beep(StdErr);
        StdErr.WriteLine("[warn] Failed to generate " + $nFailed + " of the Samurai sudokus requested.");
    }
}

function Cell(sudoku, row, column) {
    this.value = 0;
    this.blank = false;
    
    if (arguments.length) {
        this.rowIndex = row;
        this.colIndex = column;
        this.parentSudoku = sudoku;
        this.parentRow = sudoku[row];
    }
    else {
        this.rowIndex = this.colIndex = -1;
        this.parentSudoku = this.parentRow = null;
    }

    this.box = null;
    
    this.possibilities = null;
    
    this.getBox = function() {
        if (!this.box) {
            var i = this.rowIndex - this.rowIndex % 3;
            var j = this.colIndex - this.colIndex % 3;
            
            this.box = [
                this.parentSudoku[i][j],
                this.parentSudoku[i][j + 1],
                this.parentSudoku[i][j + 2],
                this.parentSudoku[i + 1][j],
                this.parentSudoku[i + 1][j + 1],
                this.parentSudoku[i + 1][j + 2],
                this.parentSudoku[i + 2][j],
                this.parentSudoku[i + 2][j + 1],
                this.parentSudoku[i + 2][j + 2]
            ];
        }
        
        return this.box;
    };
}

function newSudoku() {
    var sudoku = [];
    for (var i = 0; i < 9; i++) {
        sudoku[i] = [];
        for (var j = 0; j < 9; j++)
            sudoku[i][j] = new Cell(sudoku, i, j);
    }
    
    return sudoku;
}

function $newSudoku() {
    var $sudoku = [];
    
    for (var i = 0; i < 5; i++) {
        if (i == $CENTRAL)
            continue;
        
        $sudoku[i] = [];
        for (var j = 0; j < 9; j++) {
            $sudoku[i][j] = [];
            for (var k = 0; k < 9; k++)
                $sudoku[i][j][k] = new Cell;
        }
    }
    
    $sudoku[$CENTRAL] = [];
    for (var i = 3; i < 6; i++) {
        $sudoku[$CENTRAL][i] = [];
        for (var j = 0; j < 9; j++)
            $sudoku[$CENTRAL][i][j] = new Cell;
    }
    
    for (var i = 0; i < 3; i++) {
        $sudoku[$CENTRAL][i] = [
            $sudoku[$UPPER_LEFT][i + 6][6],
            $sudoku[$UPPER_LEFT][i + 6][7],
            $sudoku[$UPPER_LEFT][i + 6][8],
            new Cell, new Cell, new Cell,
            $sudoku[$UPPER_RIGHT][i + 6][0],
            $sudoku[$UPPER_RIGHT][i + 6][1],
            $sudoku[$UPPER_RIGHT][i + 6][2]
        ];
    }
    
    for (var i = 6; i < 9; i++) {
        $sudoku[$CENTRAL][i] = [
            $sudoku[$LOWER_LEFT][i - 6][6],
            $sudoku[$LOWER_LEFT][i - 6][7],
            $sudoku[$LOWER_LEFT][i - 6][8],
            new Cell, new Cell, new Cell,
            $sudoku[$LOWER_RIGHT][i - 6][0],
            $sudoku[$LOWER_RIGHT][i - 6][1],
            $sudoku[$LOWER_RIGHT][i - 6][2]
        ];
    }
    
    return $sudoku;
}

function fillBox(sudoku, rowStart, colStart) {
    if (sudoku[rowStart][colStart].value)
        return;
    
    var digitPool = [ 1, 2, 3, 4, 5, 6, 7, 8, 9 ];
    
    for (var i = 0; i < 3; i++) {
        for (var j = 0; j < 3; j++)
            sudoku[rowStart + i][colStart + j].value = digitPool.pick();
    }
}

function fillDiagonal(sudoku, $flip) {
    if ($flip) {
        fillBox(sudoku, 0, 6);
        fillBox(sudoku, 3, 3);
        fillBox(sudoku, 6, 0);
    }
    else {
        fillBox(sudoku, 0, 0);
        fillBox(sudoku, 3, 3);
        fillBox(sudoku, 6, 6);
    }
}

function existsInBox(box, digit) {
    for (var i = 0; i < 9; i++) {
        if (box[i].value == digit)
            return true;
    }
    
    return false;
}

function existsInRow(row, digit) {
    for (var j = 0; j < 9; j++) {
        if (row[j].value == digit)
            return true;
    }
    
    return false;
}

function existsInColumn(sudoku, column, digit) {
    for (var i = 0; i < 9; i++) {
        if (sudoku[i][column].value == digit)
            return true;
    }
    
    return false;
}

function checkIfSafe(cell, digit) {
    return (!existsInRow(cell.parentRow, digit)) &&
            (!existsInColumn(cell.parentSudoku, cell.colIndex, digit)) &&
            (!existsInBox(cell.getBox(), digit));
}

function getRemaining(sudoku) {
    var remaining = [];
    for (var i = 0; i < 9; i++) {
        for (var j = 0; j < 9; j++) {
            if (!sudoku[i][j].value)
                remaining.push(sudoku[i][j]);
        }
    }
    
    return remaining;
}

function $getRemaining($grid) {
    var $remaining = [];
    for (var i = 0; i < 9; i++) {
        for (var j = 0; j < 9; j++) {
            if (!$grid[i][j].value) {
                $grid[i][j].rowIndex = i;
                $grid[i][j].colIndex = j;
                $grid[i][j].parentRow = $grid[i];
                $grid[i][j].parentSudoku = $grid;
                $grid[i][j].box = null;
                $remaining.push($grid[i][j]);
            }
        }
    }
    
    return $remaining;
}

function fillRemaining(remaining) {
    for (var i = 0; i < 54; i++) {
        var c = remaining[i];
        c.possibilities = [];
        for (var n = 1; n <= 9; n++) {
            if (checkIfSafe(c, n))
                c.possibilities.push(n);
        }
        
        var d = c.possibilities.pick();
        if (d)
            c.value = d;
        else {
            c.possibilities = null;
            
            while (true) {
                if (!i)
                    throw new Error;
                
                var c = remaining[--i];
                var d = c.possibilities.pick();
                if (d) {
                    c.value = d;
                    break;
                }
                else {
                    c.value = 0;
                    c.possibilities = null;
                }
            }
        }
    }
}

function erasePortion(sudoku, portion) {
    switch (portion.scope) {
    case PER_SUDOKU :
    case PER_GRID :
        var unit = [];
        for (var i = 0; i < 9; i++) {
            for (var j = 0; j < 9; j++)
                unit.push(sudoku[i][j]);
        }
        
        eraseUnit();
        break;
    case PER_ROW :
        for (var i = 0; i < 9; i++) {
            var unit = [];
            
            for (var j = 0; j < 9; j++)
                unit.push(sudoku[i][j]);
            
            eraseUnit();
        }
        
        break;
    case PER_COLUMN :
        for (var j = 0; j < 9; j++) {
            var unit = [];
            
            for (var i = 0; i < 9; i++)
                unit.push(sudoku[i][j]);
            
            eraseUnit();
        }
        
        break;
    case PER_BOX :
        for (var i = 0; i < 9; i += 3) {
            for (var j = 0; j < 9; j += 3) {
                var unit = [
                    sudoku[i][j],
                    sudoku[i][j + 1],
                    sudoku[i][j + 2],
                    sudoku[i + 1][j],
                    sudoku[i + 1][j + 1],
                    sudoku[i + 1][j + 2],
                    sudoku[i + 2][j],
                    sudoku[i + 2][j + 1],
                    sudoku[i + 2][j + 2]
                ];
                eraseUnit();
            }
        }
        
        break;
    }
    
    function eraseUnit() {
        if ("size" in portion)
            var x = portion.size;
        else
            var x = Math.randomEx(portion.lBound, portion.uBound);
    
        for (var i = 0; i < x; i++)
            unit.pick().blank = true;
    }
}

function $erasePortion($sudoku, portion) {
    switch (portion.scope) {
    case PER_SUDOKU :
        var unit = [];
        for (var i = 0; i < 5; i++) {
            for (var j = 0; j < 9; j++) {
                for (var k = 0; k < 9; k++) {
                    var c = $sudoku[i][j][k];
                    if (!c.copied) {
                        unit.push(c);
                        c.copied = true;
                    }
                }
            }
        }
        
        if ("size" in portion)
            var x = portion.size;
        else
            var x = Math.randomEx(portion.lBound, portion.uBound);
    
        eraseUnit();
        break;
    case PER_GRID :
        var unit = [];
        for (var i = 0; i < 9; i++) {
            for (var j = 0; j < 9; j++) {
                $sudoku[$CENTRAL][i][j].foo = true;
                unit.push($sudoku[$CENTRAL][i][j]);
            }
        }
        
        if ("size" in portion)
            var x = portion.size;
        else
            var x = Math.randomEx(portion.lBound, portion.uBound);
        
        eraseUnit();
        var cx = x;
        
        for (var k = 0; k < 5; k++) {
            if (k == $CENTRAL)
                continue;
            
            var m = 0;
            unit.clear();
            for (var i = 0; i < 9; i++) {
                for (var j = 0; j < 9; j++) {
                    var c = $sudoku[k][i][j];
                    if (c.blank)
                        m++;
                    
                    if (!c.foo)
                        unit.push(c);
                }
            }
            
            if ("size" in portion)
                var x = portion.size;
            else
                var x = Math.randomEx(portion.lBound, portion.uBound);
            
            x -= m;
            if (x < 0)
                x = cx - m;
            
            eraseUnit();
        }
        
        break;
    case PER_BOX :
        for (var k = 0; k < 5; k++) {
            for (var i = 0; i < 9; i += 3) {
                for (var j = 0; j < 9; j += 3) {
                    var unit = [
                        $sudoku[k][i][j],
                        $sudoku[k][i][j + 1],
                        $sudoku[k][i][j + 2],
                        $sudoku[k][i + 1][j],
                        $sudoku[k][i + 1][j + 1],
                        $sudoku[k][i + 1][j + 2],
                        $sudoku[k][i + 2][j],
                        $sudoku[k][i + 2][j + 1],
                        $sudoku[k][i + 2][j + 2]
                    ];
                    
                    if (!unit[0].copied) {
                        unit[0].copied = true;
                        
                        if ("size" in portion)
                            var x = portion.size;
                        else
                            var x = Math.randomEx(portion.lBound, portion.uBound);
    
                        eraseUnit();
                    }
                }
            }
        }
        
        break;
    }
    
    function eraseUnit() {
        for (var i = 0; i < x; i++)
            unit.pick().blank = true;
    }
}

function generateSudoku(portionToErase) {
    var sudoku = newSudoku();
    
    fillDiagonal(sudoku);
    fillRemaining(getRemaining(sudoku));
    erasePortion(sudoku, portionToErase);
    
    return sudoku;
}

function $generateSudoku(portionToErase) {
    var $sudoku = $newSudoku();
    
    fillDiagonal($sudoku[$UPPER_LEFT], false);
    fillDiagonal($sudoku[$CENTRAL], false);
    fillDiagonal($sudoku[$LOWER_RIGHT], false);
    
    fillRemaining($getRemaining($sudoku[$UPPER_LEFT]));
    fillRemaining($getRemaining($sudoku[$CENTRAL]));
    fillRemaining($getRemaining($sudoku[$LOWER_RIGHT]));
    
    fillDiagonal($sudoku[$UPPER_RIGHT], true);
    fillDiagonal($sudoku[$LOWER_LEFT], true);
    
    fillRemaining($getRemaining($sudoku[$UPPER_RIGHT]));
    fillRemaining($getRemaining($sudoku[$LOWER_LEFT]));
    
    $erasePortion($sudoku, portionToErase);
    
    return $sudoku;
}

function outputRow(textStream, row) {
    for (var i = 0; i < row.length; i++) {
        if ((i != 0) && (i % 3 == 0))
            textStream.Write("|");
        
        textStream.Write(row[i].blank ? "0" : row[i].value);
    }
}

function outputSudoku(textStream, sudoku) {
    for (var i = 0; i < 9; i++) {
        outputRow(textStream, sudoku[i]);
        
        textStream.Write("\r\n");
        if ((i == 2) || (i == 5))
            textStream.WriteLine("-----------");
    }
}

function $outputSudoku(textStream, $sudoku) {
    function overlapRows(firstRow, middleRow, lastRow) {
        var result = [];
        
        for (var i = 0; i < 9; i++)
            result.push(firstRow[i]);
        
        result.push(middleRow[3], middleRow[4], middleRow[5]);
        
        for (var i = 0; i < 9; i++)
            result.push(lastRow[i]);
        
        return result;
    }
    
    for (var i = 0; i < 6; i++) {
        outputRow(textStream, $sudoku[$UPPER_LEFT][i]);
        textStream.Write("     ");
        outputRow(textStream, $sudoku[$UPPER_RIGHT][i]);
        textStream.Write("\r\n");
        
        if ((i == 2) || (i == 5))
            textStream.WriteLine("-----------     -----------");
    }
    
    for (var i = 0; i < 3; i++) {
        outputRow(textStream, overlapRows(
            $sudoku[$UPPER_LEFT][i + 6],
            $sudoku[$CENTRAL][i],
            $sudoku[$UPPER_RIGHT][i + 6]
        ));
        textStream.Write("\r\n");
    }
    
    textStream.WriteLine("        -----------        ");
    
    for (var i = 3; i < 6; i++) {
        textStream.Write("        ");
        outputRow(textStream, $sudoku[$CENTRAL][i]);
        textStream.Write("        \n");
    }
    
    textStream.WriteLine("        -----------        ");
    
    for (var i = 0; i < 3; i++) {
        outputRow(textStream, overlapRows(
            $sudoku[$LOWER_LEFT][i],
            $sudoku[$CENTRAL][i + 6],
            $sudoku[$LOWER_RIGHT][i]
        ));
        textStream.Write("\r\n");
    }
    
    textStream.WriteLine("-----------     -----------");
    
    for (var i = 3; i < 9; i++) {
        outputRow(textStream, $sudoku[$LOWER_LEFT][i]);
        textStream.Write("     ");
        outputRow(textStream, $sudoku[$LOWER_RIGHT][i]);
        textStream.Write("\r\n");
        
        if (i == 5)
            textStream.WriteLine("-----------     -----------");
    }
}

function beep(textStream) {
    textStream.Write(String.fromCharCode(7));
}