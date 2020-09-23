[APPINFO]
pGradientFill.vbp
2005 May 12
redbird77@earthlink.net
http://home.earthlink.net/~redbird77

[ABOUT]  
Yes, I propose we rename PSC to Planet-Gradient-Code!  Like an infant I am oddly fascinated by pretty colors.  In this submission I use the GradientFill API function to create horizontal and vertical repeating multicolor gradients.  I also provide a custom gradient fill function to implement smoooooth cosine and tacky HLS (rainbow) color interpolation.

No angle support, yet.  For my stab at angled gradients see CodeID=59020.

[USAGE]
To render a gradient all you need is either:

the GradientSimple function if you just need a non-repeating, two color gradient OR

the Gradient function if you need a repeating multicolor gradient.

If you want cosine color interpolation you need to include GradientFillCustom. 

[REVISIONS]
2005 May 12
	Initial release.