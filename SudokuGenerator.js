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

var wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";

var ppBorderBottom = 3;
var ppBorderLeft = 2;
var ppBorderRight = 4;
var ppBorderTop = 1;

var RGB_GREEN = 0x50b000;
var RGB_GRAY = 0xc0c0c0;

with (WSH) {
    var fso = CreateObject("Scripting.FileSystemObject");
    var wshShell = CreateObject("WScript.Shell");

    if (fso.GetBaseName(FullName).toLowerCase() != "cscript") {
        showError("This program must be run using 'cscript.exe'.");
        Quit(1);
    }

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
    
    var plainArgs = Arguments.Unnamed;
    var ts = StdOut;
    var tsIsFile = false;
    var sudokus = null, $sudokus = null;
    if (plainArgs.Count) {
        var fileName = (new Enumerator(plainArgs)).item();
        var fileExtension = fso.GetExtensionName(fileName).toLowerCase();
        switch (fileExtension) {
        case "bmp" :
        case "dib" :
        case "png" :
        case "gif" :
        case "jpg" :
        case "jpeg" :
        case "jpe" :
        case "jfif" :
        case "tif" :
        case "tiff" :
        case "wmf" :
        case "emf" :
            sudokus = [];
            $sudokus = [];
            break;
        default :
            ts = fso.OpenTextFile(fileName, 2, true);
            tsIsFile = true;
        }
    }
    
    var includeSolution = namedArgs.Exists("IS");
    var outputSolution = false;
    var nFailed = 0;
    for (var i = 0; i < n; i++) {
        try {
            var sudoku = generateSudoku(portion);
            
            if (sudokus)
                sudokus.push(sudoku);
            else {
                if (i)
                    ts.WriteBlankLines(1);
                
                outputSolution = false;
                outputSudoku(ts, sudoku);
                
                if (includeSolution) {
                    ts.WriteBlankLines(1);
                    outputSolution = true;
                    outputSudoku(ts, sudoku);
                }
            }
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
            
            if ($sudokus)
                $sudokus.push($sudoku);
            else {
                ts.WriteBlankLines(1);
                outputSolution = false;
                $outputSudoku(ts, $sudoku);
                
                if (includeSolution) {
                    ts.WriteBlankLines(1);
                    outputSolution = true;
                    $outputSudoku(ts, $sudoku);
                }
            }
        }
        catch (err) {
            StdErr.WriteLine("[warn] " + err.message);
            $nFailed++;
        }
    }
    
    if (tsIsFile)
        ts.Close();
    
    if (nFailed) {
        beep(StdErr);
        StdErr.WriteLine("[warn] Failed to generate " + nFailed + " of the normal sudokus requested.");
    }
    
    if ($nFailed) {
        beep(StdErr);
        StdErr.WriteLine("[warn] Failed to generate " + $nFailed + " of the Samurai sudokus requested.");
    }
    
    var preventMetaData = namedArgs.Exists("PMD");
    createSudokuImages();
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
    case PER_ROW :
        var unit = [];
        for (var i = 0; i < 9; i++) {
            unit.clear();
            for (var j = 0; j < 9; j++)
                unit.push($sudoku[$CENTRAL][i][j]);
            
            if ("size" in portion)
                var x = portion.size;
            else
                var x = Math.randomEx(portion.lBound, portion.uBound);
            
            $sudoku[$CENTRAL][i][2].foo = $sudoku[$CENTRAL][i][6].foo = x;
            eraseUnit();
        }
        
        eraseRows($sudoku[$UPPER_LEFT]);
        eraseRows($sudoku[$LOWER_LEFT], 1);
        eraseRows($sudoku[$UPPER_RIGHT], 0, 1);
        eraseRows($sudoku[$LOWER_RIGHT], 1, 1);
        break;
    case PER_COLUMN :
        var unit = [];
        for (var j = 0; j < 9; j++) {
            unit.clear();
            for (var i = 0; i < 9; i++)
                unit.push($sudoku[$CENTRAL][i][j]);
            
            if ("size" in portion)
                var x = portion.size;
            else
                var x = Math.randomEx(portion.lBound, portion.uBound);
            
            $sudoku[$CENTRAL][2][j].foo = $sudoku[$CENTRAL][6][j].foo = x;
            eraseUnit();
        }
        
        eraseColumns($sudoku[$UPPER_LEFT]);
        eraseColumns($sudoku[$UPPER_RIGHT], 1);
        eraseColumns($sudoku[$LOWER_LEFT], 0, 1);
        eraseColumns($sudoku[$LOWER_RIGHT], 1, 1);
        break;
    }
    
    function eraseUnit() {
        for (var i = 0; i < x; i++)
            unit.pick().blank = true;
    }
    
    function eraseRows($grid) {
        for (var i = 0; i < 9; i++) {
            unit.clear();
            if (arguments[1] ? (i > 2) : (i < 6)) {
                for (var j = 0; j < 9; j++)
                    unit.push($grid[i][j]);
                
                if ("size" in portion)
                    x = portion.size;
                else
                    x = Math.randomEx(portion.lBound, portion.uBound);
            }
            else {
                x = $grid[i][arguments[2] ? 0 : 8].foo;
                for (var j = 0; j < 9; j++) {
                    if (arguments[2] ? (j > 2) : (j < 6))
                        unit.push($grid[i][j]);
                    else if ($grid[i][j].blank)
                        x--;
                }
            }
            
            eraseUnit();
        }
    }
    
    function eraseColumns($grid) {
        for (var j = 0; j < 9; j++) {
            unit.clear();
            if (arguments[1] ? (j > 2) : (j < 6)) {
                for (var i = 0; i < 9; i++)
                    unit.push($grid[i][j]);
                
                if ("size" in portion)
                    x = portion.size;
                else
                    x = Math.randomEx(portion.lBound, portion.uBound);
            }
            else {
                x = $grid[arguments[2] ? 0 : 8][j].foo;
                for (var i = 0; i < 9; i++) {
                    if (arguments[2] ? (i > 2) : (i < 6))
                        unit.push($grid[i][j]);
                    else if ($grid[i][j].blank)
                        x--;
                }
            }
            
            eraseUnit();
        }
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
        
        textStream.Write((row[i].blank && (!outputSolution)) ? "0" : row[i].value);
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

function createSudokuImages() {
    if (!sudokus)
        return;
    
    if ((!sudokus.length) && (!$sudokus.length))
        return;
    
    fileName = fso.GetAbsolutePathName(fileName);
    
    var powerPoint = WSH.CreateObject("PowerPoint.Application");
    var presentation = powerPoint.Presentations.Add(0);
    var slide = presentation.Slides.Add(1, 12);
    
    if (!includeSolution) {
        if ((sudokus.length == 1) && ($sudokus.length == 0)) {
            createSudokuImage(sudokus[0], fileName);
            presentation.Close();
            powerPoint.Quit();
            writeMetaData();
            return;
        }
        else if (($sudokus.length == 1) && (sudokus.length == 0)) {
            $createSudokuImage($sudokus[0], fileName);
            presentation.Close();
            powerPoint.Quit();
            writeMetaData();
            return;
        }
    }
    
    var baseName = fso.GetBaseName(fileName);
    var parentFolder = fso.GetParentFolderName(fileName);
    var baseFrame = null;
    var ip = WSH.CreateObject("WIA.ImageProcess");
    var n = sudokus.length + $sudokus.length;
    for (var i = 0; i < n; i++) {
        if (includeSolution) {
            var tempFile1 = fso.BuildPath(parentFolder, baseName + "." + (i * 2) + ".png");
            var tempFile2 = fso.BuildPath(parentFolder, baseName + "." + (i * 2 + 1) + ".png");
            if (i < sudokus.length) {
                createSudokuImage(sudokus[i], tempFile1, false);
                createSudokuImage(sudokus[i], tempFile2, true);
            }
            else {
                $createSudokuImage($sudokus[i - sudokus.length], tempFile1, false);
                $createSudokuImage($sudokus[i - sudokus.length], tempFile2, true);
            }
            
            loadFrame(tempFile1);
            loadFrame(tempFile2);
        }
        else {
            var tempFile = fso.BuildPath(parentFolder, baseName + "." + i + ".png");
            if (i < sudokus.length)
                createSudokuImage(sudokus[i], tempFile, false);
            else
                $createSudokuImage($sudokus[i - sudokus.length], tempFile, false);
            
            loadFrame(tempFile);
        }
    }
    
    presentation.Close();
    powerPoint.Quit();
    
    ip.Filters.Add(ip.FilterInfos("Convert").FilterID);
    ip.Filters(ip.Filters.Count).Properties("FormatID") = wiaFormatTIFF;
    
    var tiff = ip.Apply(baseFrame);
    tiff.SaveFile(fileName);
    
    if (includeSolution)
        n *= 2;
    
    for (var i = 0; i < n; i++) {
        var tempFile = fso.BuildPath(parentFolder, baseName + "." + i + ".png");
        fso.DeleteFile(tempFile);
    }
    
    writeMetaData();
    
    function saveShapeAsImage(shape, fileName) {
        var format;
        var fileExtension = fso.GetExtensionName(fileName);
        switch (fileExtension) {
        case "bmp" :
        case "dib" :
            format = 3;
            break;
        case "png" :
            format = 2;
            break;
        case "gif" :
            format = 0;
            break;
        case "jpg" :
        case "jpeg" :
        case "jpe" :
        case "jfif" :
            format = 1;
            break;
        case "wmf" :
            format = 4;
            break;
        case "emf" :
            format = 5;
            break;
        case "tif" :
        case "tiff" :
            beep(WSH.StdErr);
            WSH.StdErr.WriteLine("[error] Saving sudokus in TIFF image format is not currently supported.");
            return;
        }
        
        shape.Export(fileName, format);
    }
    
    function createSudokuImage(sudoku, fileName, solution) {
        var table = createSudokuTable();
        
        for (var i = 0; i < 9; i++) {
            for (var j = 0; j < 9; j++) {
                var c = sudoku[i][j];
                if ((!solution) && c.blank)
                    continue;
                
                var tf = table.Cell(i + 1, j + 1).Shape.TextFrame;
                tf.HorizontalAnchor = 2;
                tf.VerticalAnchor = 3;
                tf.TextRange.Text = c.value;
                tf.TextRange.Font.Size = 24;
                if (c.blank)
                    tf.TextRange.Font.Color.RGB = RGB_GREEN;
            }
        }
        
        saveShapeAsImage(table.Parent, fileName);
        table.Parent.Delete();
        
        function createSudokuTable() {
            var table = slide.Shapes.AddTable(9, 9, -1, -1, 340, 340).Table;
            
            table.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}");
            table.TableDirection = 1;
            
            table.Columns.Item(1).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            table.Columns.Item(3).Cells.Borders.Item(ppBorderRight).Weight = 3;
            table.Columns.Item(7).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            table.Columns.Item(9).Cells.Borders.Item(ppBorderRight).Weight = 3;
            
            table.Rows.Item(1).Cells.Borders.Item(ppBorderTop).Weight = 3;
            table.Rows.Item(3).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            table.Rows.Item(7).Cells.Borders.Item(ppBorderTop).Weight = 3;
            table.Rows.Item(9).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            
            return table;
        }
    }
    
    function $createSudokuImage($sudoku, fileName, solution) {
        var $table = $createSudokuTable();
        
        for (var i = 1; i <= 6; i++) {
            for (var j = 1; j <= 21; j++) {
                if ((9 < j) && (j < 13))
                    continue;
                
                var c = (j <= 9) ? $sudoku[$UPPER_LEFT][i - 1][j - 1] :
                            $sudoku[$UPPER_RIGHT][i - 1][j - 13];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
        }
        
        for (var i = 7; i <= 9; i++) {
            for (var j = 1; j <= 6; j++) {
                var c = $sudoku[$UPPER_LEFT][i - 1][j - 1];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
            
            for (var j = 7; j <= 15; j++) {
                var c = $sudoku[$CENTRAL][i - 7][j - 7];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
            
            for (var j = 16; j <= 21; j++) {
                var c = $sudoku[$UPPER_RIGHT][i - 1][j - 13];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
        }
        
        for (var i = 10; i <= 12; i++) {
            for (var j = 7; j <= 15; j++) {
                var c = $sudoku[$CENTRAL][i - 7][j - 7];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
        }
        
        for (var i = 13; i <= 15; i++) {
            for (var j = 1; j <= 6; j++) {
                var c = $sudoku[$LOWER_LEFT][i - 13][j - 1];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
            
            for (var j = 7; j <= 15; j++) {
                var c = $sudoku[$CENTRAL][i - 7][j - 7];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
            
            for (var j = 16; j <= 21; j++) {
                var c = $sudoku[$LOWER_RIGHT][i - 13][j - 13];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
        }
        
        for (var i = 16; i <= 21; i++) {
            for (var j = 1; j <= 21; j++) {
                if ((9 < j) && (j < 13))
                    continue;
                
                var c = (j <= 9) ? $sudoku[$LOWER_LEFT][i - 13][j - 1] :
                            $sudoku[$LOWER_RIGHT][i - 13][j - 13];
                if ((!solution) && c.blank)
                    continue;
                
                fillCell($table.Cell(i, j), c);
            }
        }
        
        saveShapeAsImage($table.Parent, fileName);
        $table.Parent.Delete();
        
        function fillCell(tc, c) {
            var tf = tc.Shape.TextFrame;
            tf.HorizontalAnchor = 2;
            tf.VerticalAnchor = 3;
            tf.TextRange.Text = c.value;
            tf.TextRange.Font.Size = 18;
            if (c.blank)
                tf.TextRange.Font.Color.RGB = RGB_GREEN;
        }
        
        function $createSudokuTable() {
            var $table = slide.Shapes.AddTable(21, 21, -1, -1, 624, 624).Table;
            
            $table.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}");
            $table.TableDirection = 1;
            
            $table.Columns.Item(1).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            $table.Columns.Item(3).Cells.Borders.Item(ppBorderRight).Weight = 3;
            $table.Columns.Item(7).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            $table.Columns.Item(9).Cells.Borders.Item(ppBorderRight).Weight = 3;
            $table.Columns.Item(13).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            $table.Columns.Item(15).Cells.Borders.Item(ppBorderRight).Weight = 3;
            $table.Columns.Item(19).Cells.Borders.Item(ppBorderLeft).Weight = 3;
            $table.Columns.Item(21).Cells.Borders.Item(ppBorderRight).Weight = 3;
            
            $table.Rows.Item(1).Cells.Borders.Item(ppBorderTop).Weight = 3;
            $table.Rows.Item(3).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            $table.Rows.Item(7).Cells.Borders.Item(ppBorderTop).Weight = 3;
            $table.Rows.Item(9).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            $table.Rows.Item(13).Cells.Borders.Item(ppBorderTop).Weight = 3;
            $table.Rows.Item(15).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            $table.Rows.Item(19).Cells.Borders.Item(ppBorderTop).Weight = 3;
            $table.Rows.Item(21).Cells.Borders.Item(ppBorderBottom).Weight = 3;
            
            var col = $table.Columns.Item(10).Cells;
            for (var i = 1; i <= 6; i++)
                llg_removeBorders(col.Item(i));
            
            for (var i = 16; i <= 21; i++)
                ulg_removeBorders(col.Item(i));
            
            col = $table.Columns.Item(11).Cells;
            for (var i = 1; i <= 6; i++) {
                col.Item(i).Borders.Item(ppBorderTop).Weight = 0;
                col.Item(i).Borders.Item(ppBorderTop).Visible = 0;
                col.Item(i).Shape.Fill.BackColor.RGB = RGB_GRAY;
            }
            
            for (var i = 16; i <= 21; i++) {
                col.Item(i).Borders.Item(ppBorderBottom).Weight = 0;
                col.Item(i).Borders.Item(ppBorderBottom).Visible = 0;
                col.Item(i).Shape.Fill.BackColor.RGB = RGB_GRAY;
            }
            
            col = $table.Columns.Item(12).Cells;
            for (var i = 1; i <= 6; i++)
                lrg_removeBorders(col.Item(i));
            
            for (var i = 16; i <= 21; i++)
                urg_removeBorders(col.Item(i));
            
            col = $table.Rows.Item(10).Cells;
            for (var i = 1; i <= 6; i++)
                urg_removeBorders(col.Item(i));
            
            for (var i = 16; i <= 21; i++)
                ulg_removeBorders(col.Item(i));
            
            col = $table.Rows.Item(11).Cells;
            for (var i = 1; i <= 6; i++) {
                col.Item(i).Borders.Item(ppBorderLeft).Weight = 0;
                col.Item(i).Borders.Item(ppBorderLeft).Visible = 0;
                col.Item(i).Shape.Fill.BackColor.RGB = RGB_GRAY;
            }
            
            for (var i = 16; i <= 21; i++) {
                col.Item(i).Borders.Item(ppBorderRight).Weight = 0;
                col.Item(i).Borders.Item(ppBorderRight).Visible = 0;
                col.Item(i).Shape.Fill.BackColor.RGB = RGB_GRAY;
            }
            
            col = $table.Rows.Item(12).Cells;
            for (var i = 1; i <= 6; i++)
                lrg_removeBorders(col.Item(i));
            
            for (var i = 16; i <= 21; i++)
                llg_removeBorders(col.Item(i));
            
            return $table;
            
            function ulg_removeBorders(c) {
                c.Borders.Item(ppBorderRight).Weight = 0;
                c.Borders.Item(ppBorderBottom).Weight = 0;
                c.Borders.Item(ppBorderRight).Visible = 0;
                c.Borders.Item(ppBorderBottom).Visible = 0;
                c.Shape.Fill.BackColor.RGB = RGB_GRAY;
            }

            function urg_removeBorders(c) {
                c.Borders.Item(ppBorderLeft).Weight = 0;
                c.Borders.Item(ppBorderBottom).Weight = 0;
                c.Borders.Item(ppBorderLeft).Visible = 0;
                c.Borders.Item(ppBorderBottom).Visible = 0;
                c.Shape.Fill.BackColor.RGB = RGB_GRAY;
            }

            function llg_removeBorders(c) {
                c.Borders.Item(ppBorderTop).Weight = 0;
                c.Borders.Item(ppBorderRight).Weight = 0;
                c.Borders.Item(ppBorderTop).Visible = 0;
                c.Borders.Item(ppBorderRight).Visible = 0;
                c.Shape.Fill.BackColor.RGB = RGB_GRAY;
            }

            function lrg_removeBorders(c) {
                c.Borders.Item(ppBorderLeft).Weight = 0;
                c.Borders.Item(ppBorderTop).Weight = 0;
                c.Borders.Item(ppBorderLeft).Visible = 0;
                c.Borders.Item(ppBorderTop).Visible = 0;
                c.Shape.Fill.BackColor.RGB = RGB_GRAY;
            }
        }
    }
    
    function loadFrame(fileName) {
        var frame = WSH.CreateObject("WIA.ImageFile");
        frame.LoadFile(fileName);
        
        if (baseFrame) {
            ip.Filters.Add(ip.FilterInfos("Frame").FilterID);
            ip.Filters(ip.Filters.Count).Properties("ImageFile") = frame;
        }
        else
            baseFrame = frame;
    }
    
    function writeMetaData() {
        if (preventMetaData)
            return;
        
        if (!fso.FileExists(fileName))
            return;
        
        var n = sudokus.length;
        if (n) {
            try {
                var ts = fso.OpenTextFile(fileName + ":Sudoku", 2, true);
                
                for (var i = 0; i < n; i++) {
                    if (i)
                        ts.WriteBlankLines(1);
                    
                    outputSolution = false;
                    outputSudoku(ts, sudokus[i]);
                    
                    if (includeSolution) {
                        ts.WriteBlankLines(1);
                        outputSolution = true;
                        outputSudoku(ts, sudokus[i]);
                    }
                }
                
                ts.Close();
            }
            catch (err) {}
        }
        
        n = $sudokus.length;
        if (n) {
            try {
                var ts = fso.OpenTextFile(fileName + ":Samurai", 2, true);
                
                for (var i = 0; i < n; i++) {
                    if (i)
                        ts.WriteBlankLines(1);
                    
                    outputSolution = false;
                    $outputSudoku(ts, $sudokus[i]);
                    
                    if (includeSolution) {
                        ts.WriteBlankLines(1);
                        outputSolution = true;
                        $outputSudoku(ts, $sudokus[i]);
                    }
                }
                
                ts.Close();
            }
            catch (err) {}
        }
    }
}

function beep(textStream) {
    textStream.Write(String.fromCharCode(7));
}

function showError(message) {
    wshShell.Popup(message, 0, "Sudoku Generator", 16);
}