require 'docx'
require 'json'
require 'tk'
require 'tkextlib/tile'

class Phrase

    attr_reader :phrase_search, :phrases

    def initialize(phrases=nil,phrase_search=nil)
        @phrase_search = []
        @phrases = []
        @count = 0
    end

    def parse_phrases(docxArray)
        docxArray.each do |d| 
            doc = Docx::Document.open(d)
            document = []
            doc.each_paragraph do |pa|
                document << pa
            end
            @phrase_search << document.to_json
        end
    end

    def find_phrases(doc_files)
        fp = []
        doc_files
        @phrase_search.each_with_index do |sp,i|
            puts doc_files[i]
            @phrases.each do |ph|
                ph = ph.strip
                phrase = sp.downcase.scan(ph)
                occurrence = phrase.length
                fp.push({"count" => occurrence, "phrase" => ph, "document" => File.basename(doc_files[i]) })
            end
        end
        fp
    end

    def main
        @doc_files = []
        
        root = TkRoot.new { title "Phrase Finder" }
        content = Tk::Tile::Frame.new(root) {padding "6 6 14 0"}.grid :column => 0, :row => 0, :sticky => "nwes"
        
        TkGrid.columnconfigure root, 0, :weight => 1
        TkGrid.rowconfigure root, 0, :weight => 1

        # labels
        file_title = Tk::Tile::Label.new(content) do
            text 'Choose a .docx file'
        end

        phrase_title = Tk::Tile::Label.new(content) do
            text 'Put each phrase or word on a separate line below'
        end

        #buttons
        file_upload = Tk::Tile::Button.new(content) do
            text "Upload"
        end

        file_search = Tk::Tile::Button.new(content) do
            text "Search"
        end

        file_delete_btn = Tk::Tile::Button.new(content) do
            text "Delete Docx"
        end

        #textfields && listboxes

        file_phrases = TkText.new(content) {width 40; height 20}

        file_text = TkListbox.new(content) do
            width 20    
            height 20
            borderwidth 1
            font TkFont.new('helvetica 11')
        end

        phrase_text = TkText.new(content) do
            width 50
            height 40
            borderwidth 1
            font TkFont.new('helvetica 11')
        end

        ys = Tk::Tile::Scrollbar.new(root) {orient 'vertical'; command proc{|*args| phrase_text.yview(*args);}}
        phrase_text['yscrollcommand'] = proc{|*args| ys.set(*args);}
        ## canvas

        file_chart = TkCanvas.new(content) { width 500; height 700 }

        TkcLine.new(file_chart,50,250,50,50,'width'=>2)
        TkcLine.new(file_chart,50,250,250,250,'width'=>2)

        1.upto(10) do |i|
            x = 50 + (i*20)
            y = 250 - (i *20)
            TkcLine.new(file_chart,45,y,50,y)
            TkcText.new(file_chart,46,y,'text'=>10*i,'anchor'=>'e')
        end



        # action items on click
        file_delete = Proc.new {
            selected_item= file_text.curselection
            selected_item.each do |i|
                file_text.delete(i)
            end
        }

        file_click = Proc.new {
            doc = Tk.getOpenFile('filetypes' => "{{Docx files} {.docx}}", 'multiple' => true)
            doc = doc.split(' ')
            doc.each do |doc|
                file_text.insert 'end', File.basename(doc) + "\n"
                @doc_files << doc
            end
        }

        search_click = Proc.new{
            phrase_text.insert 'end', ""
            self.parse_phrases(@doc_files)
            @phrases = []
            file_phrases.get('1.0','end').each_line do |line|
                @phrases << line            
            end

            phrase_array = self.find_phrases(@doc_files)
            phrase_array.each do |pi|
                phrase_text.insert 'end', "Document: " + pi["document"] + "\nPhrase: " + pi["phrase"] + "\nOccurs: " + pi["count"].to_s + " times\n\n" 
            end

        }

        file_title.grid :column => 0, :row => 0, :rowspan => 2, :padx => 15, :pady => 15, :sticky => 'nsew'
        phrase_title.grid     :column => 1, :row => 1
        file_upload.grid    :column => 0, :row => 0, :rowspan => 2, :pady => 5
        file_delete_btn.grid    :column => 0, :row => 2,:padx => 5, :sticky => 'w'
        file_text.grid      :column => 0, :row => 1, :rowspan => 2, :padx => 20
        phrase_text.grid      :column => 0, :row => 3
        file_phrases.grid     :column => 1, :row => 2, :sticky => 'w', :padx => 20
        file_chart.grid      :column => 4, :row => 0, :rowspan => 4, :sticky => 'nsew', :padx => 10
        ys.grid               :column => 0, :row => 2, :sticky => 'ns'
        file_search.grid      :column => 0, :row => 3

        file_delete_btn.command = file_delete
        file_upload.command = file_click
        file_search.command = search_click

        Tk.mainloop
    end

end

phrase1 = Phrase.new()
phrase1.main





