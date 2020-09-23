Please note that this program only works with a color depth True Color 
or higher!!!

You have to admit this is by far the most original program in PSC.
Last year, i made a morphing program the likes of which had never
been seen in PSC (it can still be downloaded!). People complained that
it was too slow, and rightly so. I looked into the program structure and
found that there was a number of ways of speding it up. Here is a list of
improvements.
1)	No more triangle filling calculations.
	all the triangles are drawn using the api onto a hdc, and then
	index of the triangle is found by using their colors values.
2)	Triangle are not allowed to be overlapped.
3)	Uses DIBS to speed up program.
4)	Lots of improvements in converting Co-ortinates.
5)	Mdi environment for ease of use.
6)	All in all, a  better interface.

and not to forget...

7)	NOW YOU CAN CREATE -------->AVI<------- OF YOUR MORPHS !!!!!

Please give me some votes, and send some comments for improvements!
Thanks.