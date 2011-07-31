#!/usr/bin/ruby

#############################################################################
# Configuration
#############################################################################
# unix line breaks
# set to 1 if using doxygen on unix with
# windows formatted sources
UnixLineBreaks = 0

# leading shift inside classes/namespaces/etc.
# default is "\t" (tab)
ShiftRight     = "\t"
#ShiftRight="    "

# add namespace definition at the beginning using project directory name
# should be enabled, if no explicit namespaces are used in the sources
# but doxygen should recognize package names.
# in C# unlike in VB .NET a namespace must always be defined
leadingNamespace = 1

#############################################################################
# helper variables, don't change
#############################################################################
printedFilename    = 0
fileHeader         = 0
fullLine           = 1
insideClass        = 0
insideVB6Class     = 0
insideVB6ClassName = ""
insideVB6Header    = 0
insideSubClass     = 0
insideNamespace    = 0
insideComment      = 0
insideImports      = 0
isInherited        = 0
insideEnum         = 0
lastLine           = ""
appShift           = ""
enumComment        = ""

FILENAME = ARGV[0]

def get_filename
  FILENAME.gsub(/\\/, "/")
end

def isStatement(line)
  (line =~ /.*(private|public|protected|friend|Const|Declare)\s+/ ||
  line =~ /^(Sub|Function|Event)\s+/ ||
  line =~ /.*\s+(Sub|Function|Event|Const)\s+/)  &&
  line !~ /Exit/
end

## converts a single type definition to c#
##  "Var As Type" -> "Type Var"
def convertSimpleType(param)
  aParam = param.split(" ")
  aParam.each.with_index do |element,j|
    if element == "As"
      aParam[j]   = aParam[ j-1 ]
      aParam[j-1] = aParam[ j+1 ]
      aParam[j+1] = ""
    end
  end
  aParam.select{|e| e != ""}.join(" ")
end

def findEndArgs(string)
  string.index(/(?=[^\(])\)/) || 0
end

def isDefinition(line)
  line =~ /^(Interface|Class|Structure|Type)\s+/ ||
  line =~ /.*\s(Interface|Class|Structure)\s+/
end

def isEndDefinition(line)
  line =~ /.*End\s+(Interface|Class.*|Structure|Type)/
end

def isDefinition2(line)
  line =~ /.*(Sub|Function|Property|Event|Operator)\s.+/
end

def csharp_style_definition(line)
  line.sub!(/(?:^.*)Private\s+/  ,"private ")
  line.sub!(/(?:^.*)Public\s+/   ,"public ")

  # friend is the same as internal in c#, but Doxygen doesn't support internal,
  # so make it private to get it recognized by Doxygen) and Friend appear
  # in Documentation
  line.sub!(/(?:^.*)Friend\s+/   ,"private Friend ")
  line.sub!(/(?:^.*)Protected\s+/,"protected ")

  # add "static" to all Shared members
  line.sub!(/(?:^.*)Shared/      , "static Shared")
  
  line.sub!("Interface","interface")
  line.sub!("Class","class")
  line.sub!("Structure","struct")
  line.sub!("Type","struct")
  
  line
end

def csharp_style_param(line)
    line.gsub!("ByVal","")
    #line.gsub!("Dim","private")

    # keep ByRef to make pointers differ from others
    #gsub("ByRef","");

    # simple member definition without brackets
    if line.index("(") == nil # )
      line = convertSimpleType(line)
    elsif isDefinition2(line)
      # parametrized member
      preparams= line[0,line.index("(")]  # )

      apreparams = preparams.strip.split(" ")
      lpreparams = apreparams.length

      open_blacket = line.index("(") #)
      params  = line[open_blacket + 1, findEndArgs(line) - open_blacket -1]

      aparams = params ? params.strip.split(",") : []
      lparams = aparams.length
      
      params  = ""
      # loop over params and convert them
      if lparams > 0

        params = aparams.collect {|e|
          if e =~ /.+[()].*/
            aParam = e.split(" ")

            aParam.each.with_index do |element,j|
              if element == "As"
                aParam[j-1] = aParam[j-1].gsub(/\(\)/,"")
                aParam[j+1] = aParam[j+1].gsub(/\(\)/,"")+ "[]"
              end
            end
        
            aParam.join(" ")
          else
            e
          end
        }.collect {|e|
          convertSimpleType e
        }.join(", ")
        
        postparams = line[findEndArgs(line)+1..-1]
      else
        #postparams = line[0,rindex(line, ")")+1]
        postparams = line[line.rindex(")") + 1..-1]
      end

      #postparams=substr(line,findEndArgs(line)+1)
      # handle type def of functions and properties
      apostparams = postparams.strip.split(" ")

      if  apostparams.length > 0 && apostparams[0] == "As"
        ## functions with array as result
        apostparams[1] = apostparams[1].gsub(/\(.*\)/,"[]")

        ##
        apreparams[lpreparams]    = apreparams[lpreparams -1]
        apreparams[lpreparams -1] = apostparams[1]
        lpreparams     =+ 1
        apostparams[0] = ""
        apostparams[1] = ""
      end

      # put everything back together
      line =  apreparams.select{|e| e != nil}.join(" ")
      line =  line + "(" + params + ") "
      line =  line + apostparams.select{|e| e != nil}.join(" ")
      
    else
      # convert arrays

      line = convertSimpleType(line)

      aLine = line.split(" ")
      lLine = aLine.length

      for j in (0..lLine-1)
        if aLine[j] =~ /.*\(.*\).*/
          aLine[j]   = aLine[j].gsub(/\(.*\)/,"")
          aLine[j-1] = aLine[j-1] + "[]"
        end
      end
      line = aLine.select{|e| e != nil}.join(" ")
    end
end

exit unless File.exist? FILENAME

lines = IO.foreach FILENAME

lines.each do |line|
  line.chomp!
  #############################################################################
  # apply dos2unix
  #############################################################################
  line.sub!(/\r$/,"") if UnixLineBreaks == 1

  #############################################################################
  # merge multiline statements into one line
  #############################################################################
  if fullLine==0
    fullLine=1
    line= lastLine + line
    lastLine=""
  end

  if line =~ /_$/
    fullLine=0
    line.sub!(/_$/,"")
    lastLine=line
    next
  end

  #############################################################################
  # remove leading whitespaces and tabs
  #############################################################################
  if line =~ /^[ \t]/
    line.sub!(/^[ \t]*/, "")
  end

  #############################################################################
  # remove Option and Region statements
  #############################################################################
  if (line =~ /^#Region\s+/ || line =~ /.*Option\s+/) && insideComment != 1
    next
  end

  #############################################################################
  # VB6 file headers including class definitions
  #############################################################################

  # if file begins with a class definition, swith to VB6 mode
  if line =~ /.*\s+CLASS/ || line =~ /.*\s+VB\.Form\s+/ || line =~ /.*\s+VB\.UserControl\s+/
    insideVB6Class=1
    next
  end

  # ignore first line in VB6 forms
  if line =~ /.*VERSION\s+[0-9]+/
    next
  end

  # get VB6 class name
  if line =~ /^Attribute\s+VB_Name.*/
    insideVB6ClassName = line.gsub(/.*VB_Name\s+[=]\s+\"(.*)\"/,$1)
    insideVB6Header = 1
  end

  # detect when class attributes begin, to recognize the end of VB6 header
  if line =~ /^Attribute\s+.*/
    insideVB6Header = 1
    next
  end

  # detect the end of VB6 header
  if line !~/^Attribute\s+.*/ && insideVB6Class==1 && insideVB6Header<=1
    if insideVB6Header==0
      next
    else
      insideVB6Header = 2
    end
  end

  #############################################################################
  # parse file header comment
  #############################################################################
  if line =~ /^\s*'/ && fileHeader!=2
    # check if header already processed
    if fileHeader == 0
      fileHeader      = 1
      printedFilename = 1
      # puts @file line at the beginning
      file = get_filename
      puts "/**\n * @file " + File.basename(file)
      # if inside VB6 class module, then the file header describes the
      # class itself and should be printed after
      if insideVB6Class == 1
        puts " * \\brief Single VB6 class module, defining " + insideVB6ClassName
        puts " */"
        if leadingNamespace == 1 # leading namespace enabled?
          # get project name from the file path
          puts "namespace " + File.basename(File.dirname(file)) + " {"
          appShift = appShift + ShiftRight
        end
        puts appShift + " /**"
      end
    end
    line.sub!(/^[ \t]*'+/," * ")		# replace leading "'"
    puts appShift + line
    next
  end

  # if .*' didn't match but header was processed, then
  # the header ends here
  if fileHeader != 2
    if fileHeader != 0
      puts appShift + " */"
    end
    fileHeader = 2
  end

  #############################################################################
  # puts simply @file, if no file header found
  #############################################################################
  if printedFilename == 0

    printedFilename=1
    file = get_filename

    if insideVB6Class != 1
      puts "/// @file " + File.basename(file) + "\n"
    else
      puts "/**\n * @file " + File.basename(file)
      puts " * \\brief Single VB6 class module, defining " + insideVB6ClassName
      puts " */"
      if leadingNamespace==1 # leading namespace enabled?
        # get project name from the file path
        puts "namespace " + File.basename(File.dirname(file))  + " {"
        appShift = appShift + ShiftRight
      end
    end
  end

  #############################################################################
  # skip empty lines
  #############################################################################
  next if line =~ /^$/

  #############################################################################
  # convert Imports to C# style
  #
  # remark: doxygen seems not to recognize
  #         c# using directives so converting Imports is maybe useless?
  #############################################################################
  if line =~ /.*Imports\s+/
    line.sub!("Imports","using")
    puts line + ";"
    insideImports = 1
    next
  end

  #############################################################################
  # puts leading namespace after the using section (if presend)
  # or after the file header.
  # namespace name is extracted from file path. the last directory name in
  # the path, usually the project folder, is used.
  #
  # can be disabled by leadingNamespace=0;
  #############################################################################
  if line !~ /^Imports\s+/ && leadingNamespace <= 1 && fileHeader == 2
    if leadingNamespace == 1 # leading namespace enabled?
      # if inside VB6 file, then namespace was already printed
      if insideVB6Class != 1
        file = get_filename
        # get project name from the file path
        puts "namespace " + File.basename(File.dirname(file)) + " {"
        appShift = appShift + ShiftRight
      end
      leadingNamespace = 2	# is checked by the END function to puts corresponding "}"
    else
      # reduce leading shift
      leadingNamespace = 3
    end

    insideImports = 0

    if insideVB6Class == 1
      isInherited = 1
      puts appShift + "class " + insideVB6ClassName
    end
  end

  #############################################################################
  # handle comments
  #############################################################################

  ## beginning of comment
  if (line =~ /^\s*'''\s*/ || line =~ /^\s*'\s*[\\<][^ ].+/) && insideComment != 1
    if insideEnum == 1
      # if enum is being processed, add comment to enumComment
      # instead of printing it
      if enumComment != ""
        enumComment = enumComment + "\n" + appShift + "/**"
      else
        enumComment = appShift + "/**"
      end
    else
      # if inheritance is being processed, then add comment to lastLine
      # instead of printing it and process the end of
      # class/interface declaration

      if isInherited == 1
        isInherited = 0
        if (lastLine!="")
          puts appShift + lastLine
        end
        puts appShift + "{"
        appShift = appShift + ShiftRight
        lastLine = ""
      end
      puts appShift + "/**"
    end
    insideComment=1
  end

  ## strip leading '''
  if line =~ /^\s*'/
    if insideComment == 1
      commentString = line.sub(/^[ \t]*[']+/," * ")

      # if enum is being processed, add comment to enumComment
      # instead of printing it
      if insideEnum == 1
        enumComment = enumComment + "\n" + appShift + commentString
      else
        puts appShift + commentString
      end
      next
    end
  end

  ## end of comment
  if line !~ /^\s*'/ && insideComment==1
    # if enum is being processed, add comment to enumComment
    # instead of printing it
    if insideEnum == 1
      enumComment = enumComment + "\n" + appShift + " */"
    else
      puts appShift + " */"
    end
    insideComment = 0
  end

  #############################################################################
  # inline comments in c# style /** ... */
  #############################################################################
  # strip all commented lines, if not part of a comment block
  if line =~ /^'+/ && insideComment != 1
    next
  end

  if line =~ /.+'+/ && insideComment != 1
    line.sub!(/\s*'/," /**< \\brief ")
    line = line + " */"
  end

  #############################################################################
  # strip compiler options
  #############################################################################
  line.gsub!(/<.*> +/,"")
 
  #############################################################################
  # simple rewrites
  # vb -> c# style
  #############################################################################
  line = csharp_style_definition(line)

  #############################################################################
  # Enums
  #############################################################################
  if line =~ /^Enum\s+/ || line =~ /\s+Enum\s+/
    line.sub!("Enum","enum")
    line.sub!("+*\sAs.*","") # enums shouldn't have type definitions

    puts appShift + line + "\n" + appShift + "{"
    insideEnum   = 1
    lastEnumLine = ""
    appShift     = appShift + ShiftRight
    next
  end

  if line =~ /^[ \t]*End\s+Enum/ && insideEnum == 1 && lastEnumLine
    puts appShift  + lastEnumLine
    appShift = appShift[0,(appShift.length - ShiftRight.length)]
    puts appShift + "}"
    insideEnum   = 0
    lastEnumLine = ""
    enumComment  = ""
    next
  end

  if insideEnum == 1
    if lastEnumLine == ""
      lastEnumLine = line
      if (enumComment!="")
        puts enumComment
      end
      enumComment = ""
    else
      m =  %r[/\*\*<].match(lastEnumLine)
      if m
        commentPart    = lastEnumLine[0,m.begin(0)]
        definitionPart = lastEnumLine[0,m.begin(0) -2]
        
        if definitionPart == ""
          puts appShift + commentPart + ","
        else
          puts appShift + definitionPart + ", " + commentPart
        end
      end

      lastEnumLine = line
      # puts leading comment of next element, if present
      if enumComment != ""
        puts enumComment
      end

      enumComment = ""
    end
    next
  end

  #############################################################################
  # Declares
  #############################################################################
  if line =~ /.*Declare\s+/
    libName = line.gsub(/.+Lib\s+\"([^ ]*)\"\s.*/,$1)

    if match(line,"Alias") > 0
      aliasName = line.gsub(/.+Alias\s+\"([^ ]*)\"\s.*/," (Alias: #{$1})")
    end

    puts appShift + "/** Is imported from extern library: " + libName + aliasName + " */"
    libName   = ""
    aliasName = ""
  end

  # remove lib and alias from declares
  if line =~ /.*Lib\s+/
    line.sub!("Lib\s+[^\s]+","")
    line.sub!("Alias\s+[^\s]+","")
  end

  #############################################################################
  # types (handle As and Of)
  #############################################################################
  #if line =~ /.*\(Of (\S+)\).*/
  #  line = line.gsub(/[(]Of[ ]([^ ]+)[)]/, "<#{$1}>")
  #end
  line.gsub!(/.*\(Of (\S+)\).*/, $1 ? "<#{$1}>" : "" )

  #(/.*Function\s+/ ||
  #/.*Sub\s+/ ||
  #/.*Property\s+/ ||
  #/.*Event\s+/ ||
  #/.*Operator\s+/) &&
  if line =~ /.*As\s+/
    line = csharp_style_param(line)
  end

  #############################################################################
  # namespaces
  #############################################################################
  if line =~ /^Namespace\s+/ || line =~  /\s+Namespace\s+/
    line.sub!("Namespace","namespace")
    insideNamespace=1;
    puts appShift + line + " {";
    appShift = appShift + ShiftRight
    next
  end

  if line =~ /^.*End\s+Namespace/ && insideNamespace==1
    appShift = appShift[0,(appShift.length - ShiftRight.length)]
    puts appShift + "}"
    insideNamespace=0
    next
  end

  #############################################################################
  # interfaces, classes, structures
  #############################################################################
  if isDefinition(line)
    line = csharp_style_definition(line)

    # handle subclasses
    if insideClass==1
      insideSubClass=1
    else
      insideClass=1
    end

    # save class name for constructor handling
    if line =~ /.+class\s+([^ ]*).*/
      className = $1 || ""
    end

    isInherited = 1
    puts appShift + line
    next
  end

  # handle constructors
  if line =~ /.*Sub\s+New.*/ && className && className != ""
    className.sub!("New", "New ")
  end

  # handle inheritance
  if isInherited == 1
    if line =~ /^\s*(Inherits|Implements)\s+/
      if lastLine == ""
        line.sub!("Inherits",":")
        line.sub!("Implements",":")
        lastLine = line
      else
        line.sub!(/.*Inherits/,",")
        line.sub!(/.*Implements/,",")
        lastLine = lastLine + line
      end
    else
      isInherited = 0
      if lastLine != ""
        puts appShift + lastLine
      end
      puts appShift + "{"
      appShift = appShift + ShiftRight
      lastLine = ""
    end
  end

  if isEndDefinition(line) && (insideClass==1 || insideSubClass==1)
    if insideSubClass == 1
      insideSubClass = 0
    else
      insideClass = 0
    end
    appShift = appShift[0,(appShift.length - ShiftRight.length)]

    puts appShift + "}"
    className = ""
    next
  end

  #############################################################################
  # Replace Implements with a comment linking to the interface member,
  #   since Doxygen does not recognize members with names that differ
  #   from their corresponding interface members
  #############################################################################
  if line =~ /.+\s+Implements\s+/
    if line =~ /.*Property\s+.*/
      line = line.gsub(/(Implements)\s+(.+)$/, "/** Implements <see cref=\"#{$2}\"/> */")
    else
      line = line.gsub(/(Implements)\s+(.+)$/, "/**< Implements <see cref=\"#{$2}\"/> */")
    end
  end

  #############################################################################
  # Properties
  #############################################################################
  # skip VB6 Set/Let Property methods
  next if line =~ /.*Property\s+Set\s+/ || line =~ /.*Property\s+Let\s+/

  if line =~ /^Property\s+/ || line =~ /.*\s+Property\s+/
    line.sub!(/\(\)/,"")

    if line =~ /\(.+\)/
      line = line.gsub(/\(/,"[")
      line = line.gsub(/\)/,"]")
    else
      line = line.gsub(/\(\)/,"")
    end

    # add c# styled get/set methods
    if line =~ /ReadOnly/
      line = line + "\n" + appShift + "{ get { }; }"
    else
      line = line + "\n" + appShift + "{ get{ }; set{ }; }"
    end

    puts appShift + line
    next
  end

  if line =~ /.*Operator\s+/
    line = line.gsub(/.*Operator\s+([^ ]+)\s+/,"#{$1} operator ")
  end


  #############################################################################
  # process everything else
  #############################################################################
  if isStatement(line)

    # remove square brackets from reserved names
    # but do not match array brackets
    #  "Integer[]" is not replaced
    #  "[Stop]" is replaced by "Stop"

    line = line.gsub(/([^\[])([\\\]])/, $1 || "")
    line = line.gsub(/([\[])([^\\\]])/, $2 || "")

    # add semicolon before inline comment
    if line != ""

      commentPart    = line.index("/") ? line[line.index("/")..-1] : ""
      definitionPart = line[0,(line.index("/") || 0) - 1]

      if definitionPart != "" && commentPart != ""
        puts appShift + definitionPart  + "; " + commentPart
      else
        puts appShift + line + ";"
      end
    end
  end
end

# close file header if file contains no code
if fileHeader != 2 && fileHeader != 0
  puts " */"
end

if insideVB6Class==1
  puts ShiftRight + "}"
end

if leadingNamespace==2
  puts "}"
end
