+++
title = "Zola & Tabi - Hello, World!"
date = "2025-08-28"
description = "Setting up blogging with Zola & the Tabi theme."

[taxonomies]
tags = ["Setup", "Zola", "Tabi"]
+++

I figured I should further my programming with writing. I've written and published a few novels, but in my day-to-day
programming, I've never really jumped on the article/blogging bandwagon. I know, I know, if you want to be great at 
something you need to do the hard things as well as the easy things. The last ten years I've really dug into the parts 
of the job that get ignored or overlooked and it has catapulted my skills forward.

Writing a blog and programming articles are just the next step.

So. What was it like getting the blogging platform up?

A pain in the ass.

The first decision that had to be made was: how did I want to go about blogging?
Whatever tools and processes I used had to be simple, lightweight, quick to jump into, and with little thinking beyond 
the writing itself.

I'm mostly a backend and application developer. I didn't want to spend time fiddling with CSS, JavaScript, HTML, and on 
and on and on. I needed to focus on the writing instead of the tool itself. Better, it should be something I was already 
familiar with in some way and should be mostly a statically built site if possible. I didn't want to have to setup a server
just to get things running. Github.io pages was the obvious choice. It would host things, the pages would be in git, perfect!

[Zola](https://www.getzola.org/) met all these needs and is written in rust. Excellent! Markdown for the content made it 
even better! 

But then I wanted a kick ass theme. It had to support dark mode (the only real mode! Come on, fight me!) and it needed 
to support code blocks and ideally charts and math symbols. I looked through a few different options, duckquill, duckling,
tranquil and a few others. Sadly, they all had - bad - installation instructions. I liked those other themes, a great deal, 
but if I can't get them working without learning everything involved to the point that I could have created the theme 
myself, then why bother with it?

They failed to explain things, missed details, skipped over steps, and since I didn't already know Zola's setup, diagnosing 
the issues was a pain in the butt.

Luckily, I found [Tabi](https://github.com/welpo/tabi?tab=readme-ov-file) and the [Tabi Quickstart](https://github.com/welpo/tabi-start)
guide which made it so simple that I've been spending this time *blogging* instead of configuring. 

Of course, now I've got to get Zola, Tabi, and Github pages/workflows all working together to get things published...

Oh well, Thanks Zola and Tabi!