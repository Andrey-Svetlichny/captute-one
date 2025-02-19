## Feel free to reuse but please give me credit where possible

## absolutely needed for reliable operation
use AppleScript version "2.4"
use scripting additions

set resultLogging to true -- log non default adjustments in the script editor log
set runAllTests to true -- if false, stop checking for adjustments after the first no-defult setting is found
set resultsInCollections to false -- add adjusted and non-adjusted images to a collection
set AdjustmentLabelsInMetadata to true -- Information is added to the IPTC Getty Images fields - what and how many Adjustments are detected
set AdjustmentValuesInMetadata to false -- Information is added to the IPTC Contacts fields - adjustment values the user wants to search on
set logEveryAdjustment to false -- report every setting value in the script editor log -  very long listing - do this only for a few images

## FUNCTIONS
## if there are no selected variants, all variants in the current collection are processed
## !!! Progress updates are given at the very bottom line of the Script Editor window
## the script reports "Adjusted" or "Not_Adjusted" in "Getty Parent MEID" field
## The script finds adjustment values by getting the adjustments of the variant
## The script also finds adjustments by getting the crop, LCC Name, number of layers and number of styles of the variant
## the script reports the number of adjustments in the "Getty Original File Name" field
## the script lists the adjustment tags with non-default settings in the "Getty Personalities" field
## THE CLIPBOARD CAN NO LONGER BE CHECKED FOR NON DEFAULT ADJUSTMENTS IN CAPTURE ONE 21 OR LATER
#### The Apply Adjustments button is always enabled after a Copy Adjustments action, even if all the adjustments default
## The script either logs nothing, only non default adjustments, or every adjustment value
## The script may put some adjustment values in the IPTC Contact metadata fields.

## PURPOSE AND USE
## list all the adjustments of a variant, or a few variants
## Identify variants where all adjustments have been set to default but the variant still has the CO "Adjusted" tag
## survey adjustment values used in the catalog

## CONFIGURATION
## Accessibility should be enabled for Script Editor (might no longer be needed)

## When running this with 1000's of variants, configure as follows:
## set script editor preferences so the scripts do not log if the log window is not visible
## close the script editor log window (capturing too much log info can cause crashes)

## For faster operation with 1000's variants:
#### close the Capture One Viewer window
#### copy the selected variants to a new album if possible
#### ensure that Filter tool or a Smart Album is not working on the IPTC data fields that the script changes
#### ensure that Filter tool is not visible when the script is running

##Version Notes
## Version 7 February 27, 2022
## Removed the function that checks if the adjustments clipboard has contents
## Added roundDigits, modernised getVariantNameNum, getVariantType

set adjustedCounter to 0
set notadjustedCounter to 0

set progress description to "Setting Up"

tell application "Capture One"
    set theAppName to ("" & name)
    set currentDocRef to get current document
    set currentDocName to (get name of current document)
    tell currentDocRef to set currentCollName to (get name of current collection)
    set theSelectedVariantList to get selected variants
    if 0 = (count of theSelectedVariantList) then
        tell current document to tell current collection to set theVariantList to every variant
    else
        set theVariantList to theSelectedVariantList
    end if
end tell
set theVariantCount to get count of theVariantList
if theVariantCount = 0 then error "Did not find any variants in \"" & currentDocName & "\""

if resultLogging then
    tell application "Finder"
        set script_path to path to me
        set Script_Title to name of file script_path as text
    end tell
    log {"Script: " & Script_Title, "Document: " & currentDocName, "Collection: " & currentCollName, "Variants: " & theVariantCount}
end if

if resultsInCollections then
    set Result_ProjectName to "ScriptSearchResults"
    InitializeResultsCollection(Result_ProjectName, "NotAdjusted")
    tell application "Capture One" to copy ref2ResultAlbum to ref2NotAdjustedAlbum
    InitializeResultsCollection(Result_ProjectName, "IsAdjusted")
    tell application "Capture One" to copy ref2ResultAlbum to ref2IsAdjustedAlbum
end if


set progress description to "Evaluating "
set progress total steps to theVariantCount
set variantCounter to 0
set lastClipBoardState to false
set Mark_Start to GetTick_Now()

repeat with theVariant in theVariantList
    tell application "Capture One" to copy adjustments theVariant -- Step 1 of 2 do this early to give C1 time to react

    set variantCounter to variantCounter + 1
    set progress completed steps to variantCounter
    set Mark1 to GetTick_Now()
    set variantIsAdjusted to false
    set adjustmentList to ""
    set adjustmentListTag to ""
    set hasWBpresetAuto to false
    set hasDistCrop to false
    set countAdjustments to 0
    set theProperties to {} as list
    set therecordLabels to {} as list
    set therecordValues to {} as list

    set {theVariantName, theParentsVariantCount} to my getVariantNameNum(theVariant)
    set theImageType to getVariantsImageType(theVariant)
    tell application "Capture One" to tell parent image of theVariant to set {theCameraModel, theEXIFISO_T} to get {EXIF camera model, EXIF ISO}

    if 1 < variantCounter then
        set remainingTime to " " & roundDecimals(((Mark2 - Mark_Start) / variantCounter * (theVariantCount - variantCounter + 1)) / 60, 1) & " minutes remaining " & (round ((Mark2 - Mark_Start) / variantCounter * 1000) rounding to nearest) & "ms per variant" & " Last: " & lastTime
    else
        set remainingTime to ""
    end if
    set progress additional description to ("#" & variantCounter & " " & theVariantName & remainingTime)

    ## These are the default adjustment values before image file format and camera version is considered
    ## Mostly these are the defaults for a JPG file

    set DistortedPixelMarginPercent to 0.5
    set defaultSharpeningAmount to 0.0
    set defaultSharpeningRadius to 0.8
    set defaultSharpeningThreshold to 1.0
    set defaultNoiseReductionLuminance to 0.0
    set defaultNoiseReductionColor to 0.0
    set defaultNoiseReductionSinglePixel to 0.0
    set defaultColorProfile to "*Generic"
    set defaultFilmGrainImpact to 0.0
    set defaultFilmCurve to "Auto"
    set defaultWhiteBalancePreset to "Shot"

    ## These are the default enable settings
    set enableSharpeningAmount to true
    set enableSharpeningRadius to true
    set enableSharpeningThreshold to true
    set enableNoiseReductionLuminance to true
    set enableNoiseReductionColor to true
    set enableNoiseReductionSinglePixel to true
    set enableColorProfile to false
    set enableFilmGrainImpact to true
    set enableFilmCurve to false
    set enableWhiteBalancePreset to true



    if "JPG" = theImageType then
        ## Default settings for a JPG file
        set defaultColorProfile to {"Jpeg File neutral", "From File"}
        set enableColorProfile to true

    else if {"TIF", "TIFF"} contains theImageType then
        ## Default settings for a TIFF file
        set defaultFilmCurve to {"Tiff File neutral", "Auto"}

    else if "RW2" = theImageType then
        ## Default settings for a Panasonic RAW file
        set DistortedPixelMarginPercent to 8.34 -- 1/12 of the image -  15% of the image worst case
        set defaultSharpeningAmount to {120.0, 140.0, 160.0, 180.0, 200.0, 240.0, 280.0}
        set defaultSharpeningRadius to 1.0
        set defaultSharpeningThreshold to {1.0, 1.5}
        set defaultNoiseReductionLuminance to 50.0
        set defaultNoiseReductionColor to 50.0
        set defaultNoiseReductionSinglePixel to {0.0, 10.0, 20.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0}
        set defaultFilmCurve to {"Film Standard", "Auto"}
        set enableFilmCurve to true
        set defaultFilmGrainImpact to {0.0, 5.0, 10.0, 15.0, 20.0, 25.0}
        --set defaultColorProfile to {"Panasonic DMC-GM5 Generic", "Panasonic DMC-G5 Generic", "Panasonic DMC-GX7 Generic", "Panasonic DMC-GX1 Generic", "Panasonic DMC-GX8 Generic", "Panasonic DMC-GM1 Generic", "Panasonic DMC-G1 Generic"}
        set enableColorProfile to true
        if theCameraModel contains "GX1" then
            set defaultSharpeningRadius to {1.0, 1.299999952316}
        end if

    else if "ORF" = theImageType then
        ## Default settings for an Olympus RAW file
        set DistortedPixelMarginPercent to 8.34 -- 1/12 of the image -  15% of the image worst case
        set defaultSharpeningAmount to {120.0, 140.0, 180.0, 200.0}
        set defaultSharpeningRadius to 1.0
        set defaultNoiseReductionLuminance to 50.0
        set defaultNoiseReductionColor to 50.0
        set defaultNoiseReductionSinglePixel to {0.0, 5.0, 10.0, 20.0, 40.0, 50.0}
        -- set defaultColorProfile to "Olympus E-M1 Generic"
        set enableColorProfile to true
        set defaultFilmGrainImpact to {0.0, 7.0, 10.0}
        set defaultFilmCurve to "Auto"
        set enableFilmCurve to true

    else if "CR2" = theImageType then
        ## Default settings for an Canon RAW file
        set defaultNoiseReductionLuminance to 50.0
        set defaultNoiseReductionColor to 50.0
        set defaultFilmCurve to "Auto"
        set defaultSharpeningAmount to {150.0, 160.0}
        set defaultWhiteBalancePreset to {"Auto", "Shot"}
        if theCameraModel contains "G9" then
            set defaultSharpeningRadius to 0.800000011921
        else if theCameraModel contains "G11" then
            set defaultSharpeningRadius to 1.299999952316
        end if

    else if "ARW" = theImageType then
        ## Default settings for a Sony RAW file
        set DistortedPixelMarginPercent to 8.34 -- 1/12 of the image -  15% of the image worst case
        set defaultSharpeningAmount to 180.0
        set defaultSharpeningRadius to 0.800000011921
        set defaultNoiseReductionLuminance to 50.0
        set defaultNoiseReductionColor to 50.0
        -- set defaultColorProfile to "Sony A7S Generic"
        set enableColorProfile to true

        ## else if    add some other camera here
    end if

    tell application "Capture One" to set {{theCropHorCenter, theCropVerCenter, theCropWidth, theCropHeight}, theLCCName, countStyles, countLayers, {theImageWidth, theImageHeight}} to ¬
        get {crop, applied LCC name, (count of styles), (count of layers), (dimensions of parent image)} of theVariant

    if not (0 = countStyles) then
        set variantIsAdjusted to true
        set countAdjustments to countAdjustments + 1
        set adjustmentListTag to adjustmentListTag & "Styles;"
        if logEveryAdjustment then
            tell application "Capture One" to set theStyleList to get styles of theVariant
            set adjustmentList to adjustmentList & "Styles: " & joinListToString(theStyleList, "; ") & return
        else if resultLogging then
            set adjustmentList to adjustmentList & "Has " & countStyles & " Styles" & return
        end if
    else if logEveryAdjustment then
        set adjustmentList to adjustmentList & "No Styles" & return
    end if

    if (1 = countLayers) then
        if logEveryAdjustment then set adjustmentList to adjustmentList & "No Layers" & return
    else
        set variantIsAdjusted to true
        set countAdjustments to countAdjustments + 1
        set adjustmentListTag to adjustmentListTag & "Layers;"
        if logEveryAdjustment then
            tell application "Capture One"
                set theLayerList to get kind of layers of theVariant
                set theLayerList_T to ""
                repeat with theLayer in theLayerList
                    set theLayerList_T to theLayerList_T & (get theLayer as text) & "; "
                end repeat
                set theLayerList_T to text 1 thru -3 of theLayerList_T
            end tell
            set adjustmentList to adjustmentList & "Layers: " & theLayerList_T & return
        else if resultLogging then
            set adjustmentList to adjustmentList & "Has " & (countLayers - 1) & " Layers applied" & return
        end if
    end if

    if not ("missing value" = (get theLCCName as text)) then
        set variantIsAdjusted to true
        set countAdjustments to countAdjustments + 1
        set adjustmentListTag to adjustmentListTag & "LCC;"
        if logEveryAdjustment then
            set adjustmentList to adjustmentList & "LCC applied: " & theLCCName & return
        else if resultLogging then
            set adjustmentList to adjustmentList & "LCC applied" & return
        end if
    else if logEveryAdjustment then
        set adjustmentList to adjustmentList & "No LCC" & return
    end if

    ## There is an additional impact due to rotation which is not considered yet
    set HorizontalCrop to theImageWidth - theCropWidth
    set VerticalCrop to theImageHeight - theCropHeight
    if not ((0 ≥ HorizontalCrop) and (0 ≥ VerticalCrop)) then
        if (HorizontalCrop < (get DistortedPixelMarginPercent / 100 * theImageWidth)) and (VerticalCrop < (get DistortedPixelMarginPercent / 100 * theImageHeight)) and ¬
            (3 > AbsValue((theImageWidth / 2) - theCropHorCenter)) and ¬
            (3 > AbsValue((theImageHeight / 2) - theCropVerCenter)) ¬
                then
            if resultLogging or logEveryAdjustment then set adjustmentList to adjustmentList & "Distortion Crop: " & HorizontalCrop & "x" & VerticalCrop & return
            set hasDistCrop to true
        else
            set adjustmentListTag to adjustmentListTag & "Crop Applied;"
            set countAdjustments to countAdjustments + 1
            set variantIsAdjusted to true
            if logEveryAdjustment then
                set adjustmentList to adjustmentList & "Image Width: " & theImageWidth & "  Crop Width: " & theCropWidth & "   Image Height: " & theImageHeight & "   Crop Height: " & theCropHeight & return
            else if resultLogging then
                set adjustmentList to adjustmentList & "Cropped" & return
            end if
        end if
    else if logEveryAdjustment then
        set adjustmentList to adjustmentList & "No Crop" & return
    end if

    ## Very Fast way to get all the adjustment values with one Applescript exchange
    tell application "Capture One" to tell adjustments of theVariant to set theProperties to get properties
    tell my recordLabelsAndValues3(theProperties)
        set therecordLabels to its recordLabels
        set therecordValues to its its recordValues
    end tell
    ## run these to get the adjustmentlabel values for some other localisation
    --log therecordLabels
    --log therecordValues

    checkAdjustmentParams(false, "orientation", 0.0)
    checkAdjustmentParams(enableColorProfile, "color profile", defaultColorProfile)
    checkAdjustmentParams(false, "temperature", 0.0)
    checkAdjustmentParams(false, "tint", 0.0)
    checkAdjustmentParams(true, "rotation", 0.0)
    checkAdjustmentParams(true, "flip", "none")
    checkAdjustmentParams(enableFilmCurve, "film curve", defaultFilmCurve)
    checkAdjustmentParams(enableWhiteBalancePreset, "white balance preset", defaultWhiteBalancePreset)
    checkAdjustmentParams(true, "exposure", 0.0)
    checkAdjustmentParams(true, "brightness", 0.0)
    checkAdjustmentParams(true, "contrast", 0.0)
    checkAdjustmentParams(true, "saturation", 0.0)
    checkAdjustmentParams(true, "color balance master hue", 0.0)
    checkAdjustmentParams(true, "color balance master saturation", 0.0)
    checkAdjustmentParams(true, "color balance shadow hue", 0.0)
    checkAdjustmentParams(true, "color balance shadow saturation", 0.0)
    checkAdjustmentParams(true, "color balance shadow lightness", 0.0)
    checkAdjustmentParams(true, "color balance midtone hue", 0.0)
    checkAdjustmentParams(true, "color balance midtone saturation", 0.0)
    checkAdjustmentParams(true, "color balance midtone lightness", 0.0)
    checkAdjustmentParams(true, "color balance highlight hue", 0.0)
    checkAdjustmentParams(true, "color balance highlight saturation", 0.0)
    checkAdjustmentParams(true, "color balance highlight lightness", 0.0)
    checkAdjustmentParams(true, "level highlight rgb", 255.0)
    checkAdjustmentParams(true, "level highlight red", 255.0)
    checkAdjustmentParams(true, "level highlight green", 255.0)
    checkAdjustmentParams(true, "level highlight blue", 255.0)
    checkAdjustmentParams(true, "level target highlight rgb", 255.0)
    checkAdjustmentParams(true, "level target highlight red", 255.0)
    checkAdjustmentParams(true, "level target highlight green", 255.0)
    checkAdjustmentParams(true, "level target highlight blue", 255.0)
    checkAdjustmentParams(true, "level shadow rgb", 0.0)
    checkAdjustmentParams(true, "level shadow red", 0.0)
    checkAdjustmentParams(true, "level shadow green", 0.0)
    checkAdjustmentParams(true, "level shadow blue", 0.0)
    checkAdjustmentParams(true, "level target shadow rgb", 0.0)
    checkAdjustmentParams(true, "level target shadow red", 0.0)
    checkAdjustmentParams(true, "level target shadow green", 0.0)
    checkAdjustmentParams(true, "level target shadow blue", 0.0)
    checkAdjustmentParams(true, "level midtone rgb", 0.0)
    checkAdjustmentParams(true, "level midtone red", 0.0)
    checkAdjustmentParams(true, "level midtone green", 0.0)
    checkAdjustmentParams(true, "level midtone blue", 0.0)
    checkAdjustmentParams(true, "highlight recovery", 0.0)
    checkAdjustmentParams(true, "highlight adjustment", 0.0)
    checkAdjustmentParams(true, "shadow recovery", 0.0)
    checkAdjustmentParams(true, "black recovery", 0.0)
    checkAdjustmentParams(true, "white recovery", 0.0)
    checkAdjustmentParams(true, "clarity method", "natural")
    checkAdjustmentParams(true, "clarity amount", 0.0)
    checkAdjustmentParams(true, "clarity structure", 0.0)
    checkAdjustmentParams(enableSharpeningAmount, "sharpening amount", defaultSharpeningAmount)
    checkAdjustmentParams(enableSharpeningRadius, "sharpening radius", defaultSharpeningRadius)
    checkAdjustmentParams(enableSharpeningThreshold, "sharpening threshold", defaultSharpeningThreshold)
    checkAdjustmentParams(true, "sharpening halo suppression", 0.0)
    checkAdjustmentParams(enableNoiseReductionLuminance, "noise reduction luminance", defaultNoiseReductionLuminance)
    checkAdjustmentParams(true, "noise reduction details", 50.0)
    checkAdjustmentParams(enableNoiseReductionColor, "noise reduction color", defaultNoiseReductionColor)
    checkAdjustmentParams(enableNoiseReductionSinglePixel, "noise reduction single pixel", defaultNoiseReductionSinglePixel)
    checkAdjustmentParams(true, "film grain type", "fine")
    checkAdjustmentParams(enableFilmGrainImpact, "film grain impact", defaultFilmGrainImpact)
    checkAdjustmentParams(true, "film grain granularity", 50.0)
    checkAdjustmentParams(true, "moire amount", 0.0)
    checkAdjustmentParams(true, "moire pattern", 8)
    checkAdjustmentParams(true, "vignetting amount", 0.0)
    checkAdjustmentParams(false, "vignetting method", "elliptic on crop")
    checkAdjustmentParams(false, "dehaze amount", 0.0)


    ## Now this variant is assessed. Do something with the results
    if not variantIsAdjusted then
        set theAdjustmentTag to "Not_Adjusted"
        if resultLogging then
            log theVariantName & " has no Adjustments"
            if 0 < length of adjustmentList then log adjustmentList
        end if
        if resultsInCollections then tell application "Capture One" to tell currentDocRef to add inside ref2NotAdjustedAlbum variants {theVariant}
        set notadjustedCounter to notadjustedCounter + 1
    else
        set theAdjustmentTag to "Adjusted"
        if resultLogging then
            log theVariantName & " has Adjustments"
            log adjustmentList
        end if
        if resultsInCollections then tell application "Capture One" to tell currentDocRef to add inside ref2IsAdjustedAlbum variants {theVariant}
        set adjustedCounter to adjustedCounter + 1
    end if

    if AdjustmentValuesInMetadata then
        set theSharpeningAmount to roundDecimals((get item (get my getIndexOf2("sharpening amount", therecordLabels) as integer) of therecordValues), 1) as text
        set theSharpeningThreshold to roundDecimals((get item (get my getIndexOf2("sharpening threshold", therecordLabels) as integer) of therecordValues), 1) as text
        set theNoiseReductionSinglePixel to roundDecimals((get item (get my getIndexOf2("noise reduction single pixel", therecordLabels) as integer) of therecordValues), 1) as text
        set theFilmGrainImpact to roundDecimals((get item (get my getIndexOf2("film grain impact", therecordLabels) as integer) of therecordValues), 1) as text
        set theFilmCurve to (get item (get my getIndexOf2("film curve", therecordLabels) as integer) of therecordValues) as text
        set theColorProfile to (get item (get my getIndexOf2("color profile", therecordLabels) as integer) of therecordValues) as text
        --set theWBpreset to (get item (get my getIndexOf2("white balance preset", therecordLabels) as integer) of therecordValues)as text
        tell application "Capture One" to tell theVariant to ¬
            set {contact website, contact email, contact phone, contact country, contact postal code, contact state} to ¬
                {theSharpeningAmount, theSharpeningThreshold, theNoiseReductionSinglePixel, theFilmGrainImpact, theFilmCurve, theColorProfile}
    end if

    if AdjustmentLabelsInMetadata then
        set theVariantsCountTag to (get (get theParentsVariantCount as integer) as text)
        if runAllTests then
            set theAdjustmentCountTag to (get (get countAdjustments as integer) as text)
            if 4 < length of adjustmentListTag then set adjustmentListTag to text 1 thru -2 of adjustmentListTag
        else
            set theAdjustmentCountTag to ""
            set adjustmentListTag to ""
        end if
        tell application "Capture One" to tell theVariant to ¬
            set {Getty parent MEID, Getty original filename, Getty personalities, contact city} to ¬
                {theAdjustmentTag, theAdjustmentCountTag, adjustmentListTag, theVariantsCountTag}
    end if

    set Mark2 to GetTick_Now()
    set lastTime to (get MSduration(Mark1, Mark2) as text) & "ms"
    if resultLogging then log lastTime

end repeat
set theVariantList to {}
if resultLogging then log {"Adjusted: " & adjustedCounter, "Not Adjusted: " & notadjustedCounter}
tell application "Capture One" to select currentDocRef variants theSelectedVariantList
set theSelectedVariantList to {}

#####################   Handlers

on checkAdjustmentParams(enableFlag, theName, theTest)
    global variantIsAdjusted, therecordLabels, therecordValues, adjustmentList, resultLogging, runAllTests, countAdjustments, logEveryAdjustment, adjustmentListTag
    local theClass, theValue, thisParamAdjusted, theItem, theTestList, numDigits

    ## Skip test if set up for fast detection
    if (not runAllTests) and variantIsAdjusted then return
    if (not enableFlag) and (not logEveryAdjustment) then return

    set theClass to get class of theTest as text

    ## Get theValue associated with theName
    try
        set theValue to (get item (get getIndexOf2(theName, therecordLabels) as integer) of therecordValues)
    on error
        log "can't find: " & theName
        return missing value
    end try
    set thisParamAdjusted to false

    if theValue ≠ missing value then
        ## If theValue is missing value, then it must be the default value

        if ("real" = theClass) then
            if theValue = 0.0 then
                ## Avoiding trouble with very small values
                set theValue to 0.0
            else
                ## If theValue is a real number which is not 0 then round it
                set numDigits to 3
                set theValue to roundDigits(theValue, numDigits)
            end if
        end if

        ## Check that theValue matches theTest
        if ("integer" = theClass) then
            ## Integers
            if not (theTest = (get theValue as integer)) then set thisParamAdjusted to true
        else if ("real" = theClass) then
            ## Real Numbers
            if not ((get (get theTest as real) as text) = (get (get theValue as real) as text)) then set thisParamAdjusted to true
        else if ("text" = theClass) then
            if "*" = (get first item of theTest) then
                ## Text beginning with "*" --> contains
                if not ((get theValue as text) contains (get {items 2 thru -1 of theTest} as string)) then set thisParamAdjusted to true
            else
                ## Text
                if not (theTest = (get theValue as text)) then set thisParamAdjusted to true
            end if
        else if ("list" = theClass) then
            ## Lists, treated as lists of text strings
            set theTestList to {}
            repeat with theItem in theTest
                set theTestList to theTestList & (get theItem as text)
            end repeat
            if not ((get theTestList as list) contains (get theValue as text)) then set thisParamAdjusted to true
        else
            log "Unexpected class " & theClass
            return missing value
        end if
    end if

    if enableFlag and thisParamAdjusted then
        if resultLogging then set adjustmentList to adjustmentList & theName & ": " & theValue & return
        set adjustmentListTag to adjustmentListTag & theName & ";"
        set variantIsAdjusted to true
        set countAdjustments to countAdjustments + 1
    end if
    --if logEveryAdjustment then log {theName, theValue, enableFlag, thisParamAdjusted, variantIsAdjusted, countAdjustments}

    if logEveryAdjustment then log {theName, theValue, thisParamAdjusted}

    return (not thisParamAdjusted)
end checkAdjustmentParams

#####################   Capture One Utility Handlers

on getVariantNameNum(theVariant)
    tell application "Capture One" to tell theVariant to set {theVariantPosition, theBaseName} to {position, name}
    tell application "Capture One" to tell (get parent image of theVariant) to set theParentsVariantCount to (count of variants)
    if (1 = theParentsVariantCount) then return {theBaseName, theParentsVariantCount}
    return {(theBaseName & "-" & theVariantPosition), theParentsVariantCount}
end getVariantNameNum

on getVariantsImageType(theVariant)
    tell application "Capture One" to tell theVariant to set theFileExt to extension of parent image
    return theFileExt
end getVariantsImageType

on InitializeResultsCollection(nameResultProject, nameResultAlbumRoot)
    ## General Purpose Handler for scripts using Capture One Pro
    ## Sets up a project and albums for collecting images

    global ref2ResultAlbum, enableNotifications, currentDocRef

    tell application "Capture One" to tell currentDocRef
        if not (exists collection named (get nameResultProject)) then
            set ref2ResultProject to make new collection with properties {kind:project, name:nameResultProject}
        else
            if ("project" = (kind of (get collection named nameResultProject)) as text) then
                set ref2ResultProject to collection named nameResultProject
            else
                error ("A collection named \"" & nameResultProject & "\" already exists, and it is not a project.")
            end if
        end if
    end tell

    set coll_ctr to 1
    set nameResultAlbum to nameResultAlbumRoot & "_" & (get short date string of (get current date)) & "_"
    repeat
        tell application "Capture One" to tell ref2ResultProject
            if not (exists collection named (get nameResultAlbum & coll_ctr)) then
                set nameResultAlbum to (get nameResultAlbum & coll_ctr)
                set ref2ResultAlbum to make new collection with properties {kind:album, name:nameResultAlbum}
                exit repeat
            else
                set coll_ctr to coll_ctr + 1
            end if
        end tell
    end repeat

end InitializeResultsCollection

#####################   General Utility Handlers

on splitStringToList(theString, theDelim)
    ## Public Domain
    set astid to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to theDelim
        set theList to text items of theString
    on error
        set AppleScript's text item delimiters to astid
    end try
    set AppleScript's text item delimiters to astid
    return theList
end splitStringToList

to joinListToString(theList, theDelim)
    ## Public Domain
    set theString to ""
    set astid to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to theDelim
        set theString to theList as string
    on error
        set AppleScript's text item delimiters to astid
    end try
    set AppleScript's text item delimiters to astid
    return theString
end joinListToString

on removeLeadingTrailingSpaces(theString)
    ## Public Domain, modified
    repeat while theString begins with space
        -- When the string is only 1 character long, then it is exactly 1 space, and the next operation willl crash. So return ""
        if 1 ≥ (count of theString) then return ""
        set theString to text 2 thru -1 of theString
    end repeat
    repeat while theString ends with space
        set theString to text 1 thru -2 of theString
    end repeat
    return theString
end removeLeadingTrailingSpaces

on getIndexOf2(theItem, theList)
    set theCount to 0
    set theItem_T to get theItem as text
    repeat with anItem in theList
        set theCount to theCount + 1
        if (get anItem as text) = theItem_T then return theCount
    end repeat
    return 0
end getIndexOf2

on AbsValue(theValue)
    if theValue < 0 then return (get -theValue)
    return theValue
end AbsValue

on roundToQuantum(thisValue, quantum)
    ## Public domain author unknown
    return (round (thisValue / quantum) rounding to nearest) * quantum
end roundToQuantum

on roundDecimals(n, numDecimals)
    ## Nigel Garvey, Macscripter
    set x to 10 ^ numDecimals
    tell n * x to return (it div 0.5 - it div 1) / x
end roundDecimals

on roundDigits(thisValue, numDigits)
    ## Eric Valk, 2022
    set theDigits to (length of (get "" & (thisValue div 1)))
    if thisValue < 0 then
        set theDigits to theDigits - 1
    else if 1 > thisValue then
        set theDigits to theDigits - 1
    end if
    if theDigits > numDigits then
        set theQuantum to 1
    else
        set theQuantum to 10 ^ (theDigits - numDigits)
    end if
    return roundToQuantum(thisValue, theQuantum)
end roundDigits

on GetTick_Now()
    ## From MacScripter Author "Jean.O.matiC"
    ## returns duration in seconds since since 00:00 January 2nd, 2000 GMT, calculated using computer ticks
    script GetTick
        property parent : a reference to current application
        use framework "Foundation" --> for more precise timing calculations
        on Now()
            return (current application's NSDate's timeIntervalSinceReferenceDate) as real
        end Now
    end script
    return GetTick's Now()
end GetTick_Now

on MSduration(firstTicks, lastTicks)
    ## Public domain
    ## returns duration in ms
    ## inputs are durations, in seconds, from GetTick's Now()
    return (round (10000 * (lastTicks - firstTicks)) rounding to nearest) / 10
end MSduration

## Don't touch this handler code. It is Harum Scarum sophisticated
on recordLabelsAndValues3(theRecord)
    -- obtained from http://macscripter.net/viewtopic.php?id=45430
    -- the 2017 final version, very sophisticated script, From "bmose"
    -- Depends on using the clipboard
    -- Returns the unpiped and piped forms of a record's labels and the text and value forms of a record's values
    -- Utility properties and handlers

    script util
        property tokenChar : character id 60000 -- obscure character chosen for the low likelihood of its appearance in an input record's text representation; this may be substituted by any character or character string that does not appear in the text representation of the input record
        property tokenizedStrings : {}
        on detokenizeString(tokenizedString)
            -- Convert any tokens of the form "[token char][token index number][token char]" to their original values; handle nested tokens with recursive handler calls
            set tid to AppleScript's text item delimiters
            try
                set AppleScript's text item delimiters to my tokenChar
                tell (get tokenizedString's text items)
                    if length < 3 then
                        set originalString to tokenizedString
                    else
                        set originalString to (item 1) & my detokenizeString((my tokenizedStrings's item ((item 2) as integer)) & (items 3 thru -1))
                    end if
                end tell
            end try
            set AppleScript's text item delimiters to tid
            return originalString
        end detokenizeString
        on representValueAsText(theValue)
            -- Parse a forced error message for the text representation of the input value
            try
                || of {theValue}
            on error m
                try
                    if m does not contain "{" then error
                    repeat while m does not start with "{"
                        set m to m's text 2 thru -1
                    end repeat
                    if m does not contain "}" then error
                    repeat while m does not end with "}"
                        set m to m's text 1 thru -2
                    end repeat
                    if m = "{}" then error
                    set valueAsText to m's text 2 thru -2
                on error
                    try
                        -- Try an alternative method of generating a text representation from a forced error message if the first method fails
                        {||:{theValue}} as null
                    on error m
                        try
                            if m does not contain "{" then error
                            repeat while m does not start with "{"
                                set m to m's text 2 thru -1
                            end repeat
                            set m to m's text 2 thru -1
                            repeat while m does not start with "{"
                                set m to m's text 2 thru -1
                            end repeat
                            if m does not contain "}" then error
                            repeat while m does not end with "}"
                                set m to m's text 1 thru -2
                            end repeat
                            set m to m's text 1 thru -2
                            repeat while m does not end with "}"
                                set m to m's text 1 thru -2
                            end repeat
                            if m = "{}" then error
                            set valueAsText to m's text 2 thru -2
                        on error
                            error "Can't get a text representation of the value."
                        end try
                    end try
                end try
            end try
            return valueAsText
        end representValueAsText
    end script
    -- Perform the handler's actions inside a try block to capture any errors; if an error is encountered, restore AppleScript's text item delimiters to their baseline value
    set tid to AppleScript's text item delimiters
    try
        -- Handle the special case of an empty record
        if theRecord = {} then return {recordLabels:{}, recordLabelsPiped:{}, recordValuesAsText:{}, recordValues:{}}
        -- Get the text representation of the input record
        set textValue to util's representValueAsText(theRecord)
        -- Partially validate the text representation, and remove the leading and trailing curly braces
        tell textValue
            -- This test does not exclude an input argument in the form of an Applescript list; however, a list will result in a parsing error in the code below
            if (it does not start with "{") or (it does not end with "}") then error "The input value is not a record."
            set textValue to text 2 thru -2
        end tell
        -- Initialize return values and the token counter
        set {recordLabels, recordLabelsPiped, recordValuesAsText, recordValues, iToken} to {{}, {}, {}, {}, 0}
        -- Tokenize text elements that could potentially contain record property (", ") or label/value (":") delimiter characters in order to avoid errors while parsing record properties
        -- Tokens are of the form "[token char][token index number][token char]", where the index numbers increase sequentially 1, 2, 3, ... with each new token
        -- Tokenize escaped double-quote characters (to facilitate tokenizing double-quoted items)
        set AppleScript's text item delimiters to "\\\""
        tell (get textValue's text items)
            if length > 1 then
                set iToken to iToken + 1
                set AppleScript's text item delimiters to (util's tokenChar) & iToken & (util's tokenChar)
                set {textValue, end of util's tokenizedStrings} to {it as text, "\\\""}
            end if
        end tell
        -- Tokenize double-quoted items
        set AppleScript's text item delimiters to "\""
        tell (get textValue's text items)
            if length > 2 then
                set textValue to ""
                repeat with i from 1 to (length - 2) by 2
                    set iToken to iToken + 1
                    set {textValue, end of util's tokenizedStrings} to {textValue & (item i) & (util's tokenChar) & iToken & (util's tokenChar), "\"" & item (i + 1) & "\""}
                end repeat
                set textValue to textValue & item -1
            end if
        end tell
        -- Tokenize piped items
        set AppleScript's text item delimiters to "|"
        tell (get textValue's text items)
            if length > 2 then
                set textValue to ""
                repeat with i from 1 to (length - 2) by 2
                    set iToken to iToken + 1
                    set {textValue, end of util's tokenizedStrings} to {textValue & (item i) & (util's tokenChar) & iToken & (util's tokenChar), "|" & item (i + 1) & "|"}
                end repeat
                set textValue to textValue & item -1
            end if
        end tell
        -- Tokenize curly-braced items
        set AppleScript's text item delimiters to "{"
        tell (get textValue's text items)
            if length > 1 then
                set {AppleScript's text item delimiters, textValue, iNestedLevel, currBracedString} to {"}", item 1, 1, ""}
                repeat with s in rest
                    tell (get s's text items) to if length > 1 then set {t, iNestedLevel} to {it, iNestedLevel + 1 - length}
                    if iNestedLevel > 0 then
                        set {iNestedLevel, currBracedString} to {iNestedLevel + 1, currBracedString & s & "{"}
                    else
                        set iToken to iToken + 1
                        set {textValue, iNestedLevel, currBracedString, end of util's tokenizedStrings} to {textValue & (util's tokenChar) & iToken & (util's tokenChar) & t's item -1, 1, "", "{" & currBracedString & (t's items 1 thru -2) & "}"}
                    end if
                end repeat
            end if
        end tell
        -- Get the unpiped and piped record labels and text representations of the record values
        set AppleScript's text item delimiters to ", " -- delimits record properties
        tell (get textValue's text items)
            set AppleScript's text item delimiters to ":" -- delimits the label and value for a given record property
            repeat with s in it
                -- Extract and detokenize the current record label and value
                tell (get s's text items) to set {end of recordLabelsPiped, end of recordValuesAsText} to {util's detokenizeString(item 1), util's detokenizeString(item 2)}
                -- Create an unpiped version of the current record label
                tell recordLabelsPiped's item -1
                    if it = "||" then
                        set end of recordLabels to ""
                    else if (it starts with "|") and (it ends with "|") then
                        set end of recordLabels to text 2 thru -2
                    else
                        set end of recordLabels to it
                    end if
                end tell
            end repeat
        end tell
        -- Get the record values
        set recordValues to theRecord as list
    on error m number n
        set AppleScript's text item delimiters to tid
        if n = -128 then error number -128
        if n ≠ -2700 then set m to "(" & n & ") " & m -- -2700 = purposely thrown error
        error "Handler recordLabelsAndValues error:" & return & return & m
    end try
    set AppleScript's text item delimiters to tid
    -- Return the results
    return {recordLabels:recordLabels, recordLabelsPiped:recordLabelsPiped, recordValuesAsText:recordValuesAsText, recordValues:recordValues}
end recordLabelsAndValues3

