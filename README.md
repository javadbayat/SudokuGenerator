# A command-line Sudoku Generator
Generates a Sudoku puzzle (either a **normal** sudoku or a **Samurai** sudoku) and outputs it either as **plain-text** or as **a raster image**.

An example of a **normal** sudoku generated using this tool (represented as **plain-text**):

```
005|460|918
910|000|302
082|031|005
-----------
093|005|207
200|000|190
871|309|004
-----------
129|050|406
007|006|831
308|107|009
```

An example of a **Samurai** sudoku generated using this tool (represented as **plain-text**):

```
103|605|208     007|902|000
457|108|096     890|005|700
600|490|070     641|873|905
-----------     -----------
700|080|409     900|520|478
240|009|001     400|030|009
809|004|652     765|409|031
-----------     -----------
502|900|803|074|150|704|302
300|807|020|180|000|000|004
904|002|007|003|084|301|007
        -----------
        600|500|002
        080|360|005
        005|907|408
        -----------
400|060|090|700|501|000|900
300|748|050|430|807|596|340
008|005|700|051|600|800|005
-----------     -----------
892|001|075     759|000|612
030|920|680     430|105|000
600|084|920     008|907|403
-----------     -----------
206|410|007     906|018|530
740|006|010     010|700|000
910|803|264     084|600|020
```

An example of a **normal** sudoku generated using this tool (represented as **a PNG image**):

![An example of a normal sudoku](Sudoku.png)

An example of a **Samurai** sudoku generated using this tool (represented as **a PNG image**):

![An example of a Samurai sudoku](Samurai.png)

## Command-line Syntax

    cscript SudokuGenerator.js [/NS:NumSudokus] [/NG:NumSamuraiSudokus]
    [/EP:RangeOfBlanks | NumBlanks [ScopeSuffix]] [/IS] [/PMD] [OutputFileName]

## Description of parameters
All parameters are optional. They're explained below.

- `/NS`: Set this parameter to the number of **normal sudokus** to be generated. By default, only one normal sudoku is generated.
- `/NG`: This parameter is required if you wish to generate **Samurai** sudokus. Set it to the number of **Samurai sudokus** to be generated. By default, no Saumrai sudokus are generated.
- `/EP`: Typically, setting this parameter is neccessary because, by default, the generated sudokus would be filled completely and contain no blank cells. Thus, to generate "usable" sudokus, you must explicitly specify the amount of cells to be erased using this parameter. It can can be simply set to an integer value indicating the number of blank cells in each generated sudoku. Alternatively, you can set this parameter to a range, like `22-44`. In this case, for each of the sudokus to be generated, the system will generate a random number within the specified range, and that random number will then determine the number of blank cells in that sudoku. Moreover, you can add a **scope suffix** at the end of this parameter's value so the specified amount of blank cells will be considered within smaller regions of the sudoku rather than within the entire sudoku. For example, if you specify `6R`, the program will make 6 blank cells within each row of the sudoku. The following scope suffixes are supported. Note that they must be specefied in upper-case, and there must be no space between the integer/range value and the suffix (e.g. `7 R` is not allowed).
    + `B`: There will be *k* blank cells **per each 3x3 block**. In this case, *k* must be at most 9.
    + `R`: There will be *k* blank cells **per each row**. In this case, *k* must be at most 9.
    + `C`: There will be *k* blank cells **per each column**. In this case, *k* must be at most 9.
    + `G`: For **Samurai sudokus**, the program will make *k* blank cells **per each overlapping 9x9 grid**. For normal sudokus, setting this suffix has no effect, and there will be *k* blank cells within the entire sudoku. For technical reasons, in case of specifying the `G` suffix, *k* must be **at most 72**.
- `/IS`: If this parameter is set, each of the sudokus in the output will be followed by its corresponding solution (answer).
- `OutputFileName`: By default, the program will output the generated sudokus as **plain-text** to **the Standard Output Stream**. If you want to store the sudokus in a file, you can **specify a file name** on the command-line. If the specified file name has one of **the extensions `.bmp`, `.dib`, `.png`, `.gif`, `.jpg`, `.jpeg`, `.jpe`, `.jfif`, `.tif`, `.tiff`, `.wmf`, or `.emf`**, then the sudokus will be stored **as raster image**. Otherwise, they will be stored **as plain-text**.

### Saving the output sudokus as image files
When saving the generated sudokus as image files, take the following into account:

If the `/IS` parameter is not set and a single sudoku is requested (either a normal sudoku or a Samurai sudoku), then the extension of the specified output file name must be either `.bmp`, `.dib`, `.png`, `.gif`, `.jpg`, `.jpeg`, `.jpe`, `.jfif`, `.wmf`, or `.emf`.

If either `/IS` parameter is set or the normal and Samurai sudokus requested in total is more than one, then the extension of the output file name must be certainly `.tif` or `.tiff`. In that case, **a multi-page TIFF file** will be created, and each of the sudokus generated will be placed in an individual page of the TIFF file.

> [!NOTE]
> For the program to be able to save the generated sudokus as image files, you **must have Microsoft Office PowerPoint installed**. In addition, in cases where a multi-page TIFF file needs to be created, your computer **must be running Windows Vista or higher**, or the operation fails.