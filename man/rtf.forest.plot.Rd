%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  forest.plot.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{rtf.forest.plot}
\alias{rtf.forest.plot}


 \title{Get an RTF encoded forest plot}

 \description{
	Get an RTF encoded forest plot.
 }

 \usage{
rtf.forest.plot(x=1.25, min=0.5, max=2, xlim=c(0.1, 12), width=3, height=0.3, cex=1,
  lwd=0.75, res=300)
}

 \arguments{
 	\item{x}{x (e.g. hazard ratio).}
 	\item{min}{Minimum whisker (e.g. lower bound of 95\% hazard ratio CI).}
 	\item{max}{Maximum whisker (e.g. upper bound of 95\% hazard ratio CI).}
 	\item{xlim}{A \code{\link[base]{vector}} specifying the x limits.}
 	\item{width}{Plot width in inches.}
 	\item{height}{Plot height in inches.}
 	\item{cex}{A numerical value giving the amount by which plotting text and symbols should be magnified relative to the default.}
 	\item{lwd}{Line width.}
 	\item{res}{Output resolution in dots per inch.}
 }

 \details{
 	Create a forest plot and convert PNG to RTF code.  This is useful for
 	embedding into a data frame of hazard ratios and then writing an
 	RTF output file.  See the example below for usage.
 }

 \examples{
 \dontrun{
 tab<-data.frame(
	Label=c("Test1","Test2","Test3"),
 	HR=c(1,2,0.45),
 	Lower.CI=c(0.5,1.1,0.25),
 	Upper.CI=c(2,3.5,0.9),
 	stringsAsFactors=FALSE,
 	check.names=FALSE)

 # create forest plots by row
 forest.plot.args<-list(xlim=c(0.1,5),width=3.0,height=0.3,cex=1,lwd=0.75,res=300)
 tab$"HR Plot (log scale)"<-mapply(rtf.forest.plot,tab$HR,tab$Lower.CI,tab$Upper.CI,
			MoreArgs=forest.plot.args)

 # rbind the x-scale to the table in the plot column
 xscale<-rtf.forest.plot.xscale(xlim=c(0.1,5),width=3.0,height=0.3,cex=1,
 			lwd=0.75,res=300)

 tab<-data.frame(lapply(tab, as.character),
 			stringsAsFactors=FALSE,
			check.names=FALSE)

 tab<-rbind(tab,list("","","","",xscale))

 # write the RTF output
 rtf<-RTF("test_rtf.forest.plot.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addTable(rtf,tab,col.widths=c(0.75,0.75,0.75,0.75,3))
 done(rtf)
 }
 }

