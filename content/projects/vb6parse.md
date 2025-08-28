+++
title = "VB6Parse"
description = "A library for parsing Visual Basic 6 projects, modules, classes, forms, and workspaces, written in rust"
weight = 1
template = "page.html"

[taxonomies]
tags = ["rust", "Visual Basic 6", "vb6"]
+++

[vb6parse](https://github.com/scriptandcompile/vb6parse) was started with the goal of building an API to allow me to read in some of the legacy projects my company has
and hopefully, one day, convert them into something else with modern tools and systems.

The requirements for vb6parse are pretty broad since I've got multiple end goals in mind.

* It needs to be a fast API that can be used in multiple other tools.
* Has to be able to handle everything involved with a standard VB6 project. Projects, Modules, Classes, Forms, Project Groups, everything.
* The library should handle multiple 'levels' of parsing. Outputting a simple AST is not enough. 
It should be able to handle raw text from a string, output a token stream, an AST, and save back to disk in a format that
Microsoft's Visual Basic 6 IDE can recognize.
* Protect the API's user from making boneheaded type mistakes if possible. I wanted to avoid the use of stringly-typed 
API's that are so common when one language interfaces with another. Idealy, I should be able to create every bit of code 
at the AST level and get a fully Visual Basic 6 project/project group saved to disk, or work the other direction and get 
a full project/project group data structure with AST's.
* Code comments and structure matter. Some parser libraries I've worked with simply throw out any kind of structure or 
comments, assuming that the API's user doesn't want them. Writing a code formatter is not easy when the parser library
hides the code's format and comments from you.
* Warning and errors need to be front and center. A tool that can't explain where things have gone wrong and instead simply
throws its hands up in the air with 'oopsie!' is not an API I can work with.

vb6parse has (mostly) made progress on every front here. 
