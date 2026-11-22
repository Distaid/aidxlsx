# Aid.Xlsx.Formula

This project parses Excel formulas and references in A1 and R1C1 notation. The workbook project
supplies an implementation of `IAstFactory` to turn parsed expressions into calculation nodes.
Calculation and worksheet functions live in the main project. The main NuGet package includes
the formula assembly.

| Directory | Responsibility |
| --- | --- |
| [Syntax](Syntax/) | Formula parser, parser errors, operators, and the AST factory contract |
| [Lexing](Lexing/) | Tokens, checked-in lexer tables in `Automata/`, and their definitions in `Grammar/` |
| [References](References/) | Cell, range, and structured references |
| [References/Transforms](References/Transforms/) | Reference rewriting for copies and A1/R1C1 conversion |

Start with [`FormulaParser`](Syntax/FormulaParser.cs) or [`ReferenceParser`](References/ReferenceParser.cs).
The project's namespaces use the `Aid.Xlsx.Formulas` root. Attribution and licensing for the parser
are documented in the repository. [NOTICE](../../NOTICE) · [License](../../third-party/LICENSE-Formula.txt)
