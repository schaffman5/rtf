%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addTrellisObject.RTF}
\alias{addTrellisObject.RTF}
\alias{RTF.addTrellisObject}
\alias{addTrellisObject,RTF-method}

 \title{Insert a trellis plot object into the RTF document}

 \description{
	Insert a trellis plot object into the RTF document.
 }

 \usage{
\method{addTrellisObject}{RTF}(this, trellis.object, width=3, height=0.3, res=300, rotate=NULL, ...)
}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{trellis.object}{The trellis plot object.}
 	\item{width}{Plot output width in inches.}
 	\item{height}{Plot output height in inches.}
 	\item{res}{Output resolution in dots per inch.}
 	\item{rotate}{Object rotation in degrees. \bold{optional}.}
 	\item{...}{Not used.}
 }

 \details{
 	Plots are added to the document as PNG objects.  Multi-page trellis objects are
	automatically split across multiple pages in the RTF output file.  To rotate the
   object to landscape orientation within the RTF output, use rotate=90.  When using
   rotation, width and height still refer to the unrotated plot dimensions and not the
   rotated output dimensions on the RTF page.  An alternative to rotating the plot is
   to rotate the entire page using a call to addPageBreak with suitable page width and
   height dimensions.
 }

 \examples{
 rtf<-RTF("test_addTrellisObject.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 if(require(lattice) & require(grid)) {
 	# multipage trellis object
 	p2<-xyplot(uptake ~ conc | Plant, CO2, layout = c(2,2))
 	addTrellisObject(rtf,trellis.object=p2,width=8,height=4,res=300, rotate=90)
 }
 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
   For more information see \code{\link{RTF}}.
 }


\keyword{methods}
