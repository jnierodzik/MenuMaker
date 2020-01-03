--
--	Created by: James Nierodzik
--	Created on: 12/15/19
--
--	Copyright © 2019 Casa de Santiago, All Rights Reserved
--   Free for personal use, but must be licenses for commercial use.
--
--   If you like this buy me a dram or better yet a bottle ;-)
--   https://paypal.me/jnierodzik
--

use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

-- read in file
set strWhiskeyFile to choose file with prompt "Choose a tab-delimited text file"
set strDataFile to paragraphs of (read strWhiskeyFile) --better than return
set {strRecordCount, whiskeyList} to {count strDataFile, {}}
tell (a reference to my text item delimiters) to set {oldDelimiter, contents} to {contents, tab}
repeat with this_row in strDataFile
	set end of whiskeyList to (text items of (contents of this_row))
end repeat
set my text item delimiters to oldDelimiter

-- setup variables
try
	--setup menu frame column bounds
	set col1MenuFrameStartingBounds to {2.427331679688, 3.95, 3.413610839844, 24}
	set col2MenuFrameStartingBounds to {2.427331679688, 29.45, 3.413610839844, 49.5}
	set menuFrameBounds to {0, 0, 0, 0}
	set initialMenuFrameSpacer to 3.186389160156
	set menuFrameSpacer to 1.6
	
	--setup fonts
	set h1Font to "Baskerville	Bold"
	set h1FontSize to 25
	set brandFont to "Baskerville	Bold"
	set brandFontSize to 13
	set expressionFont to "Baskerville	Regular"
	set expressionFontSize to 13
	set detailsFont to "Baskerville	Italic"
	set detailsFontSize to 10
	set legendFont to "Baskerville	Regular"
	set legendFontSize to 8
	
	-- setup misc variables
	set menuItemsPerPage to 16
	set newPage to true
	set newColumn to false
	set menuCounter to 0
	set pageCounter to 0
	set columnCounter to 0
	set geoSpace to 3.43334
	
	--setup column mappings in whiskey file
	set typeColumn to 8
	set brandColumn to 4
	set expressionColumn to 5
	set companyColumn to 3
	set proofColumn to 16
	set bottledInBondColumn to 9
	set barrelProofColumn to 10
	set singleBarrelColumn to 11
	set wheatedColumn to 14
	set finishedColumn to 15
	set classificationColumn to 19
	set finishingNotes to 22
	set ageColumn to 17
	set noteColumn to 6
	
	--setup layout: column 1 geometric bounds
	set col1TitleY1 to 1.5
	set col1TitleX1 to 1.5
	set col1TitleY2 to 4.3
	set col1TitleX2 to 24.0
	
	--setup layout: column 2 geometric bounds
	set col2TitleY1 to 1.5
	set col2TitleX1 to 27.0
	set col2TitleY2 to 4.3
	set col2TitleX2 to 49.5
	
	--setup layout: legend
	set col1LegendBIBy1 to 63.599998168942
	set col1LegendBIBx1 to 4.799999084471
	set col1LegendBIBy2 to 64.5
	set col1LegendBIBx2 to 5.700000915529
	
	set col1LegendBPy1 to 63.599998168942
	set col1LegendBPx1 to 12.525
	set col1LegendBPy2 to 64.5
	set col1LegendBPx2 to 13.425001831058
	
	set col1LegendWy1 to 63.599998168942
	set col1LegendWx1 to 19.35
	set col1LegendWy2 to 64.5
	set col1LegendWx2 to 20.250001831058
	
	set col1LegendBIBy1TF to 63.599998168942
	set col1LegendBIBx1TF to 6
	set col1LegendBIBy2TF to 64.5
	set col1LegendBIBx2TF to 10.4
	
	set col1LegendBPy1TF to 63.599998168942
	set col1LegendBPx1TF to 13.725
	set col1LegendBPy2TF to 64.5
	set col1LegendBPx2TF to 17.225
	
	set col1LegendWy1TF to 63.599998168942
	set col1LegendWx1TF to 20.550000915529
	set col1LegendWy2TF to 64.5
	set col1LegendWx2TF to 23.150000915529
	
	set col2LegendBIBy1 to 63.599998168942
	set col2LegendBIBx1 to 30.299999084471
	set col2LegendBIBy2 to 64.5
	set col2LegendBIBx2 to 31.200000915529
	
	set col2LegendBPy1 to 63.599998168942
	set col2LegendBPx1 to 38.025
	set col2LegendBPy2 to 64.5
	set col2LegendBPx2 to 38.925001831058
	
	set col2LegendWy1 to 63.599998168942
	set col2LegendWx1 to 44.85
	set col2LegendWy2 to 64.5
	set col2LegendWx2 to 45.750001831058
	
	set col2LegendBIBy1TF to 63.599998168942
	set col2LegendBIBx1TF to 31.5
	set col2LegendBIBy2TF to 64.5
	set col2LegendBIBx2TF to 35.9
	
	set col2LegendBPy1TF to 63.599998168942
	set col2LegendBPx1TF to 39.225
	set col2LegendBPy2TF to 64.5
	set col2LegendBPx2TF to 42.725
	
	set col2LegendWy1TF to 63.599998168942
	set col2LegendWx1TF to 46.050000915529
	set col2LegendWy2TF to 64.5
	set col2LegendWx2TF to 48.650000915529
	
	--setup image stuff
	set bibLarge to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:bib.tif"
	set bpLarge to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:barrel_proof.tif"
	set wheatedLarge to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:wheat.tif"
	set bibLegend to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:bib_legend.tif"
	set bpLegend to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:barrel_proof_legend.tif"
	set wheatedLegend to alias "Macintosh HD:Users:jamesnierodzik:Documents:Casa de Santiago:image resources:wheat_legend.tif"
	set iconSpacer to 0.127221167969
	--
end try

set currentClassification to item classificationColumn of item 2 of whiskeyList

tell application "Adobe InDesign CC 2019"
	--loop through whiskeys
	repeat with a from 2 to length of whiskeyList
		set thisWhiskey to item a of whiskeyList
		
		-- fake loop to test for duplicates (if this whiskey is the same as the previous) and skip 
		repeat 1 times
			try
				if (item expressionColumn of thisWhiskey) = (item expressionColumn of item (a - 1) of whiskeyList) then
					if (item brandColumn of thisWhiskey) = (item brandColumn of item (a - 1) of whiskeyList) then
						if (item noteColumn of thisWhiskey) = (item noteColumn of item (a - 1) of whiskeyList) then
							exit repeat
						end if
					end if
				end if
			end try
			
			-- start a new page / column if we're on a new type
			if (item classificationColumn of thisWhiskey) is not equal to currentClassification then
				set currentClassification to (item classificationColumn of thisWhiskey)
				set menuCounter to 0
				if columnCounter mod 2 is 0 then
					set newPage to true
				else
					set newColumn to true
				end if
			end if
			
			-- make a new document/page if needed
			if newPage then
				if pageCounter = 0 then
					-- Menu3 Preset information
					-- Units: Picas
					-- Width: 51p0
					-- Height: 66p0
					-- Pages 1
					-- Facing Pages: On
					-- Columns: 1
					-- Column Gutter: 1p0
					-- Margins (all): 1p6
					set myDoc to make new document with properties {document preset:"Menu3"}
					set facing pages of document preferences of myDoc to false
				else
					tell myDoc
						make page at after page pageCounter
					end tell
				end if
				set newPage to false
				set pageCounter to pageCounter + 1
				set myPage to page pageCounter of myDoc
				
				-- setup h1 frame
				my insertSimpleTextFrame(col1TitleY1, col1TitleX1, col1TitleY2, col1TitleX2, item classificationColumn of thisWhiskey, h1FontSize, h1Font, myPage)
				
				-- make legend
				my insertImage(col1LegendBIBy1, col1LegendBIBx1, col1LegendBIBy2, col1LegendBIBx2, bibLegend, myPage)
				my insertSimpleTextFrame(col1LegendBIBy1TF, col1LegendBIBx1TF, col1LegendBIBy2TF, col1LegendBIBx2TF, "Bottled In Bond", legendFontSize, legendFont, myPage)
				
				my insertImage(col1LegendBPy1, col1LegendBPx1, col1LegendBPy2, col1LegendBPx2, bpLegend, myPage)
				my insertSimpleTextFrame(col1LegendBPy1TF, col1LegendBPx1TF, col1LegendBPy2TF, col1LegendBPx2TF, "Barrel Proof", legendFontSize, legendFont, myPage)
				
				my insertImage(col1LegendWy1, col1LegendWx1, col1LegendWy2, col1LegendWx2, wheatedLegend, myPage)
				my insertSimpleTextFrame(col1LegendWy1TF, col1LegendWx1TF, col1LegendWy2TF, col1LegendWx2TF, "Wheated", legendFontSize, legendFont, myPage)
				
				set columnCounter to columnCounter + 1
				set item 1 of menuFrameBounds to item 1 of col1MenuFrameStartingBounds
				set item 2 of menuFrameBounds to item 2 of col1MenuFrameStartingBounds
				set item 3 of menuFrameBounds to item 3 of col1MenuFrameStartingBounds
				set item 4 of menuFrameBounds to item 4 of col1MenuFrameStartingBounds
			end if
			
			
			if newColumn then
				-- setup h1 frame
				my insertSimpleTextFrame(col2TitleY1, col2TitleX1, col2TitleY2, col2TitleX2, item classificationColumn of thisWhiskey, h1FontSize, h1Font, myPage)
				
				-- make legend
				my insertImage(col2LegendBIBy1, col2LegendBIBx1, col2LegendBIBy2, col2LegendBIBx2, bibLegend, myPage)
				my insertSimpleTextFrame(col2LegendBIBy1TF, col2LegendBIBx1TF, col2LegendBIBy2TF, col2LegendBIBx2TF, "Bottled In Bond", legendFontSize, legendFont, myPage)
				
				my insertImage(col2LegendBPy1, col2LegendBPx1, col2LegendBPy2, col2LegendBPx2, bpLegend, myPage)
				my insertSimpleTextFrame(col2LegendBPy1TF, col2LegendBPx1TF, col2LegendBPy2TF, col2LegendBPx2TF, "Barrel Proof", legendFontSize, legendFont, myPage)
				
				my insertImage(col2LegendWy1, col2LegendWx1, col2LegendWy2, col2LegendWx2, wheatedLegend, myPage)
				my insertSimpleTextFrame(col2LegendWy1TF, col2LegendWx1TF, col2LegendWy2TF, col2LegendWx2TF, "Wheated", legendFontSize, legendFont, myPage)
				
				set columnCounter to columnCounter + 1
				set newColumn to false
				set item 1 of menuFrameBounds to item 1 of col2MenuFrameStartingBounds
				set item 2 of menuFrameBounds to item 2 of col2MenuFrameStartingBounds
				set item 3 of menuFrameBounds to item 3 of col2MenuFrameStartingBounds
				set item 4 of menuFrameBounds to item 4 of col2MenuFrameStartingBounds
			end if
			
			
			--make a menu frame and populate it
			tell myPage
				
				-- setup menu frame bounds
				if menuCounter = 0 then -- first item on a page, space from title
					set item 1 of menuFrameBounds to ((item 1 of menuFrameBounds) + initialMenuFrameSpacer)
					set item 3 of menuFrameBounds to ((item 1 of menuFrameBounds) + initialMenuFrameSpacer)
				else -- additional item on page, space from previous
					set item 1 of menuFrameBounds to lastFrameHolder + menuFrameSpacer
					set item 3 of menuFrameBounds to ((item 1 of menuFrameBounds) + initialMenuFrameSpacer)
				end if
				
				
				set menuFrame to make text frame with properties {geometric bounds:menuFrameBounds}
				tell menuFrame
					-- insert brand
					set contents of last insertion point to item brandColumn of thisWhiskey
					tell last paragraph of menuFrame
						set point size to brandFontSize
						set applied font to brandFont
						set brandCount to count of characters
					end tell
					
					-- insert expression
					set contents of last insertion point to " " & (item expressionColumn of thisWhiskey)
					tell last paragraph of menuFrame
						set applied font of characters (brandCount + 2) thru -1 to expressionFont
					end tell
					
					-- setup details line
					set theDetails to return & (item companyColumn of thisWhiskey)
					if (item noteColumn of thisWhiskey) is not equal to "" then
						set theDetails to theDetails & " | " & (item noteColumn of thisWhiskey)
					end if
					set theDetails to theDetails & " | " & (item proofColumn of thisWhiskey) & " proof"
					if (item ageColumn of thisWhiskey) is not equal to "" then
						set theDetails to theDetails & " | " & (item ageColumn of thisWhiskey) & "yr"
					end if
					set contents of last insertion point to theDetails
					
					tell last paragraph of menuFrame
						set applied font to detailsFont
						set point size to detailsFontSize
					end tell
					
					if (item finishingNotes of thisWhiskey) is not equal to "" then
						set theDetails to return & (item finishingNotes of thisWhiskey)
						set contents of last insertion point to theDetails
					end if
					
					set auto sizing reference point of text frame preferences to top center point
					set auto sizing type of text frame preferences to height only
					
					set lastFrameHolder to item 3 of geometric bounds
					
				end tell
				set menuCounter to menuCounter + 1
			end tell
			
			-- add icons
			if columnCounter mod 2 is 0 then
				set geo1 to item 1 of geometric bounds of menuFrame
				set geo2 to 27.0
				set geo3 to (item 3 of geometric bounds of menuFrame) + iconSpacer
				set geo4 to 29.1
			else
				set geo1 to item 1 of geometric bounds of menuFrame
				set geo2 to 1.5
				set geo3 to (item 3 of geometric bounds of menuFrame) + iconSpacer
				set geo4 to 3.6
			end if
			
			-- -- Bottled in Bond
			if item bottledInBondColumn of thisWhiskey = "Y" then my insertImage(geo1, geo2, geo3, geo4, bibLarge, myPage)
			
			-- -- Barrel Proof
			if item barrelProofColumn of thisWhiskey = "Y" then my insertImage(geo1, geo2, geo3, geo4, bpLarge, myPage)
			
			-- -- Wheated
			if item wheatedColumn of thisWhiskey = "Y" then my insertImage(geo1, geo2, geo3, geo4, wheatedLarge, myPage)
			
			-- check for list size and toggle new colum/page indicator
			if menuCounter = menuItemsPerPage then
				set menuCounter to 0
				if columnCounter mod 2 is 0 then
					set newPage to true
				else
					set newColumn to true
				end if
			end if
		end repeat
	end repeat
end tell

on insertImage(y1, x1, y2, x2, imagePath, pageReference)
	tell application "Adobe InDesign CC 2019"
		tell pageReference
			set myRect to make rectangle with properties {geometric bounds:{y1, x1, y2, x2}, stroke weight:0.0}
			place imagePath on myRect
		end tell
	end tell
end insertImage

on insertSimpleTextFrame(y1, x1, y2, x2, theString, pointSize, fontType, pageReference)
	tell application "Adobe InDesign CC 2019"
		tell pageReference
			set myText to make text frame with properties {geometric bounds:{y1, x1, y2, x2}}
			tell myText
				set contents of last insertion point to theString
				tell last paragraph
					set point size to pointSize
					set applied font to fontType
				end tell
			end tell
		end tell
	end tell
end insertSimpleTextFrame
