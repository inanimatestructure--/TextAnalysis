require 'docx'
require 'json'
require 'tk'
require 'tkextlib/tile'

class Phrase

    attr_reader :phrase_search, :phrases

    def initialize(phrases=nil,phrase_search=nil)
        @phrase_search = phrase_search
        @phrases = ["north", "this is a test", "nothin for me", "wind"] 
        @count = 0
    end

    def parse_phrases
        doc = Docx::Document.open('test.docx')
        document = []
        doc.each_paragraph do |pa|
            document << pa
        end
        @phrase_search = document.to_json
    end

    def find_phrases
        @phrases.each do |ph|
            phrase = @phrase_search.downcase.scan(ph)
            occurrence = phrase.length
            puts "#{ph} occurs #{occurrence} times"
        end
    end

    def main 
        self.parse_phrases
        self.find_phrases
        @doc_files = ""
        
        root = TkRoot.new { title "Phrase Finder" }
        frame = Tk::Tile::Frame
        TkLabel.new(root) do
            text 'Choose a .docx file'
            pack { side 'left' }
        end

        file_upload = TkButton.new(root) do
            text "Upload"
            pack("side" => "top", "padx"=> "50", "pady"=> "50")
        end

        file_text = TkText.new(root) do
            width 20    
            height 10
            borderwidth 1
            font TkFont.new('helvetica 11')
            pack("side" => "right", "padx"=> "5", "pady"=> "5")
        end

        file_click = Proc.new {
            @doc_files = Tk.getOpenFile('filetypes' => "{{Docx files} {.docx}}", 'multiple' => true)
            @doc_files = @doc_files.split(' ')
            @doc_files.each do |doc|
                file_text.insert 'end', File.basename(doc) + "\n"
            end
        }

        phrase_text = TkText.new(root) do
            width 30
            height 30
            borderwidth 1
            font TkFont.new('helvetica 11')
            pack("side" => "bottom",  "padx"=> "5", "pady"=> "5")
        end
        
        file_upload.command = file_click

        Tk.mainloop
    end

end

phrase1 = Phrase.new()
phrase1.main





