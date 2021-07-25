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
        content = Tk::Tile::Frame.new(root) {padding "5 5 12 0"}.grid :column => 0, :row => 0, :sticky => "nwes"
        
        TkGrid.columnconfigure root, 0, :weight => 1
        TkGrid.rowconfigure root, 0, :weight => 1

        file_title = Tk::Tile::Label.new(content) do
            text 'Choose a .docx file'
        end

        file_upload = Tk::Tile::Button.new(content) do
            text "Upload"
        end

        file_delete = Proc.new {
            file_text.delete
        }

        file_text = TkListbox.new(content) do
            width 20    
            height 10
            borderwidth 1
            font TkFont.new('helvetica 11')
        end

        file_click = Proc.new {
            @doc_files = Tk.getOpenFile('filetypes' => "{{Docx files} {.docx}}", 'multiple' => true)
            @doc_files = @doc_files.split(' ')
            @doc_files.each do |doc|
                file_text.insert 'end', File.basename(doc) + "\n"
            end
        }

        phrase_text = TkText.new(content) do
            width 40
            height 20
            borderwidth 1
            font TkFont.new('helvetica 11')
        end

        file_title.grid :column => 0, :row => 0, :padx => 15, :pady => 15, :sticky => 'nsew'
        file_upload.grid    :column => 0, :row => 0, :pady => 5
        file_text.grid      :column => 0, :row => 2, :sticky => 'w', :padx => 20
        phrase_text.grid      :column => 0, :row => 3, :sticky => 'w', :padx => 20

        
        file_upload.command = file_click

        Tk.mainloop
    end

end

phrase1 = Phrase.new()
phrase1.main





