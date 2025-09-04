+++
title = "VB6Parse - Part 1"
date = "2025-08-29"
description = "Building a high performance VB6 parsing library - Project Design and Goals."

[taxonomies]
tags = ["Design", "Goals", "Research", "VB6", "Visual Basic 6"]
+++

One of the first things I try to emphasize when designing a project for my work is 'what problem are we trying to solve'?

It seems kind of silly, after all, why do all this work if we don't know what we want to do in the first place? 
Unfortunately, in the initial stage of building something it's very easy to allow goal creep. More than once, I've seen
new 'what if we...' and 'it would be awesome if we could do...' push things so far that the original goal gets shoved 
out of sight.

Worse, one of the most important things when designing a project is to pick which goals the project is *not* going to solve.

What I *want* is a parsing library for Visual Basic 6 in rust. We have a lot of legacy code and we do a lot of contract 
work on legacy VB6 code. It would be very nice to be able to migrate from this no longer maintained language and move on 
to more modern systems. Beyond just converting in a batch process from one language to another, we also would like to take 
full control of testing systems and even move on to a more modern IDE and compiler setup where possible. All of this should 
be possible with a decent VB6 parsing library.

Now, we definitely have a hierarchy of goals here. Some are far more important than others. I care a great deal more about
the batch process of converting a VB6 application over developing a compiler or an IDE. One small project to use our 
library that would be instantly useful is something to parse our project, check for missing dll's, ocx's, projects, and 
source code files. Another useful program would be a code formatter; a common VB6 issue is to have your variables shifting 
constantly between uppercase and lowercase. A tool that quickly formats the code based on some standard configuration 
format, à la cargo fmt, would be absolutely amazing.

Aspen is my answer to those problems. Think of aspen as the cargo of the vb6 world. It's still a rough work in progress.
It's useful to try and use your libraries as they are developed - and rust encourages splitting projects into libraries 
and the application/s that use them. Aspen allows me to experiment with usability and ergonomics when working with the
vb6parse library. 

So, what should we focus on first?

Research!

Heading to Github I grabbed a bunch of repositories which contained mostly VB6 code. After downloading a good fifty 
projects, I started reading and experimenting. The first thing that I discovered jumped out at me with a scream of horror.

Some of the downloaded code wasn't in English.

## Encoding ##

This shouldn't have been a surprise, but I hadn't really given it much of a thought at first. 
Worse, VB6 comes from before the era of Unicode and the miracle that is UTF-8/16/32. So why is this such an issue? 
After all, that there are other languages in the world isn't exactly a huge surprise. What *is* a huge surprise
is that compiling a Chinese language VB6 project fails if the country/culture/language selection in Windows is set to 
English! The problem is that since the source-code defaults to whatever Windows Code Page the OS is set to rather than 
being specified somewhere in the project file. This means we have to just *know* somehow what language the code was 
originally written in. If you want to compile some random code you might have to change your language/region/culture
in Windows and maybe even download language packs and reboot before you can compile!

Since rust strings are UTF-8, we need to first load any of these files check to see if the code conforms to the latin-1
code page (ie, is likely English) and then convert it to a UTF-8 string. This pre-pass step is a bit annoying since it
means we need to reserve a buffer for the code, fill this buffer, then create a second buffer and copy and convert from 
this first buffer into the second buffer.

Luckily for us, someone else has done most of this work for us, namely with the (encoding_rs)[https://crates.io/crates/encoding_rs] 
library. On the usability side, I want to be able to read a file by just giving the library a file path, or the contents 
of a file. A file path because this is likely the most common way to use the API, and the second because testing sucks
when you have to write a file for unit tests. Also, a parser that requires a file system is a pain when you want to write
an interpreter / or code hosting engine. I'm not sure if those are something I plan to do, but since it is required for 
ease of testing anyway, I should include it. Especially since the file based API entry point will likely just use the 
underlying string fed API entry anyway.

This should be as easy as:

```rust
use encoding_rs::{mem::utf8_latin1_up_to, CoderResult, WINDOWS_1252};

fn encode(source_code: &[u8]) -> Result<String, ErrorDetails> {
    let mut decoder = WINDOWS_1252.new_decoder();

    let Some(max_len) = decoder.max_utf8_buffer_length(source_code.len()) else {
        // This is mostly just hand waving-ism for the error reporting
        let error_offset = 0;
        let decode_len = 0;
        let details = ErrorDetails { error_offset, decoded_len };
        return details;
    }

    let mut file_content = String::Default();

    let last = true;
    let (coder_result, attempted_decode_len, all_processed) =
        decoder.decode_to_string(source_code, &mut source_file.file_content, last);

    if file_content.len() == source_code.len() {
        // looks like we actually succeeded even if the coder_result might be
        // confused at that.
        return Ok(source_file);
    }

    if (!all_processed && !allow_replacement) || coder_result == CoderResult::OutputFull {
        let mut decoded_len = utf8_latin1_up_to(source_code);
        let mut error_offset = decoded_len - 1;

        // Looks like we actually succeeded even if the coder_result might be
        // confused at that.
        if attempted_decode_len == decoded_len {
            return Ok(source_file);
        }

        let text_up_to_error = if let Ok(v) = str::from_utf8(&source_code[0..decoded_len]) {
            v.to_owned()
        } else {
            // For some reason, even though this should never happen
            // we ended up here. Oh well. Report that things failed at
            // the start of the file since we can't pinpoint the exact
            // location.
            error_offset = 0;
            decoded_len = 0;
        };

        // This is mostly just hand waving-ism for the error reporting
        let details = ErrorDetails { error_offset, decoded_len };

        return Err(details);
    }

    Ok(source_file)
}
```

This isn't as clean and simple as I would like. But, it's effective. I just need to keep in mind that the encoding_rs crate 
was designed to be used within a browser for handling decoding and encoding in a web context. It does more than it needs
to do, but that's fine since it does what I need it to do well enough.

So encoding/decoding is solved. Wonderful. Detecting when we have something other than English and reporting it is easy
enough and should give us somewhere to go if someone cares to write non-english parsing and code page recognition. It's not
a trivial problem since chinese character can get mixed in with english characters when creating variables and that whole 
thing just ends up being an entire cluster-f. I'm not interested in doing that work. It's officially 'not supported' and
I feel more than happy with an error message that explains the issue of non-English non-support.

## Pseudo-Formats ##

Next on the 'old-tech is problematic' train is the pseudo-INI format for the project / project group format. The format
*mostly* resembles an INI file (ie, name=value pairs with bracketed named sections and an unbracketed default section) 
except most INI parsers have issues with the 'names' part of that parse being anything but normal ASCII characters. When
the parser goes the extra mile and allows unicode, they usually balk at non-letter characters...

Which makes this abomination fail in most INI parsers:

```
FavorPentiumPro(tm)=0
```

Yeah. 'FavorPentiumPro(tm)=0'. That causes...more than a few issues in a few different parsers. There are a few other 
issues like certain sections having values which are double quoted, while other values have compound values (un-quoted)
that also need to be unpacked. This is annoying, but something we can handle. [vb6parse](https://github.com/scriptandcompile/vb6parse/blob/master/src/parsers/project.rs#L448) handles this well enough, but I built this part of the library first and I made some serious 
mistakes and this will need to be fixed. What I *should* have done was create a tokenizer and then parse these tokens with
a project specific parser. 

The way to handle things is a bit tricky since, while the ini-ish file is line-by-line, certain settings are incompatible
and there are some rather complex hierarchial data results which need to be correctly applied based on previous (or future!)
data. One goal that I want with this parser is that we can't create incompatible settings for the data structures the parsing
returns, but we *also* want to support byte-to-byte consistency. If the project has a DLL base address, yet the project is 
an exe, we still want to save the project and have a DLL base address value inside the file!

## Parsing ##

On the parsing side, I've gone through three phases and a bit more research at the start would have made this simpler.
I started with [nom](https://crates.io/crates/nom) which is a parser combinator library. I switched to [winnow](https://crates.io/crates/winnow)
which is a fork of nom with a lot of ergonomic changes. I originally chose nom because I wanted to learn about parser libraries
and see what it was like using them. Sadly, the problem with parser combinator libraries is that the combinator part. The very
thing that makes them so powerful is the same thing that annoyed the heck out of me while using them. Combining the smaller 
parsers to create larger systems looked great in toy examples. So clean. So pretty!

Then I tried to add errors and warning reporting. This is where things fell apart. Badly.

Suddenly those five line clean parser combination examples exploded into a huge amount of lines where I had to constantly
check for errors before I could 'combine' the parsers again. I already had that same thing with Pratt parsers and custom
parsing processes, but now I had a weird type signature and an extra level of mental abstraction in the type system I had
to keep in mind.

Nope, I decided to write a simplified parser library and rewrite it to use the new library. I'm currently in the middle
of this rewrite and it is going well enough and it's making sense. But, like all rewrites, it seems to be a quagmire of
slow, unimpressive progress which suddenly 'shifts' all at once into success.

Such is programming.

## Tokenize, *then* Parse ##

Yeah, that would have been the right way to handle it. But of course, I got all excited to replace winnow with my own custom
parser that I skipped the tokenize step and tried to go from characters to parsing. It mostly works since the parser is
so much less difficult to mentally handle, the sudden lowering of my mental-overhead from switching out the parser-combinator
which I was still learning and going for one of the more standard parsers I was already familiar with meant that I got much
further before I realized...um...maybe I shouldn't be making this mistake?

Yet another thing that needs to be fixed.

## Errors - They are very important ##

As I said before. Reporting errors is a vital part of the process. I want this tool to be very helpful when it comes to
processing VB6. There are a lot of ways things can be done wrong - as in any language - but the language has enough redundancy
that we can actually provide a lot of helpful errors to correct things. That leaves us with three kinds of results.

1) Successful parsing without any errors.
2) Successful parsing, but we have one ore more errors where we recovered or could correct for the issue.
3) Failed parsing. Whatever went wrong was so overwhelming, so terrible that we just can't make any progress.

Now result number one is easy to handle. It's exactly what we want so it's easy to plan to handle. The second result is
actually a bit trickier and annoying to handle; I would even argue that recovering from and working around syntax issues
and mistakes of that nature are the real issue for industry ready parser / compiler / interpreter. The third is a curious
case since, by definition, it's where we have failed to handle the problem and have just thrown up our hands and failed.

We could even think of failing to parse the source code as failure to manage the second case! We should think of these
three things more holistically, specifically, we should consider the first case to be the situation only where everything
person. In all other cases, wherever it is possible, we should try to move the third case into the second case wherever
possible and only fall back to the third case when we have no choice since there is just not enough information to go on.

So how should we structure these errors? My answer? They should be in two parts because they should solve two main problems.
We need to both to be able to handle - programmatically - the problem of how a parse fails. A tool that is trying to 
reformat and correct code, a parse failure is a signal where something should be fixed and how. In an interpreter or 
compiler then most of the time these errors are simply collected and reported in a way that a user can understand and deal
with. After all, if the parser understood what went wrong and it made sense how the format was slightly off, then it would
be re-written to correct things.

A good example is the 'header' language at the start of VB6 module, class, and forms. These use a language and should very 
clearly be structured correctly when saved, but reading it we should be far more forgiving. The IDE assumes newlines, 
spaces, comments, version information, and so onn and so on to be in *exactly* the right way. The IDE only accepts 
specific version numbers - likely as a protection towards backwards compatibility - but was never actually used! The whole 
version line could be removed and everything should keep working for us. We want to be forgiving in reading and strict on
writing.

For formatting the error I use [ariadne](https://crates.io/crates/ariadne) for creating and handling error messages, I had
considered [miette](https://crates.io/crates/miette) but finally settled on ariadne. I use [thiserror](https://crates.io/crates/thiserror). 
On the other hand, in [aspen](https://github.com/scriptandcompile/aspen) I use [anyhow](https://crates.io/crates/anyhow) 
since it works great within binaries while thiserror is great for libraries and they both work together well.

## Reading the file != data representation ##

There is a significant difference between reading a file and structuring that data for the purpose of parsing the language
versus using it within an interpreter or compiler. A good example of this is that in both VB6 project files and VB6 forms,
these files make references to other files rather than containing the data itself. This happens with VB6 form resources
like BMP's, listbox items, and even references to OCX files and other projects in the case of project group files.

In the parser, it would be better to just load the references and move on, storing these references for the API's user
to handle things as needed. Unfortunately, I didn't consider this while designing things and just switched over to handling
things directly since I was considering how I would use these same data structures in a tool for formatting and project 
checking.

Yet another thing that needs to be better handled in the future. Pulling in these files should mean we have two different
formats for these items. Either the data value directly, or a references to where the resource resides that contains this
data. A bit clunkier for users of the API if they want to handle this in an interpreter or what have you, but it makes the
parsing significantly easier which is the main goal of the project.



