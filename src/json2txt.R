# 必要なパッケージのインストールと読み込み
if (!require("jsonlite")) install.packages("jsonlite")
if (!require("officer")) install.packages("officer")

library(jsonlite)
library(officer)
library(tidyverse)


# JSONファイルが保存されているディレクトリを指定
json_directory <- "data/J-TREE Space Slack export Apr 20 2023 - Feb 6 2024/self-intro" # JSONファイルのディレクトリを指定

# outputのパスを指定
output_directory <- "output/J-TREE Space Slack export Apr 20 2023 - Feb 6 2024"

# ディレクトリが存在しない場合にのみ、ディレクトリを作成
if (!dir.exists(output_directory)) {
  dir.create(output_directory)
  cat("ディレクトリが正常に作成されました: ", output_directory, "\n")
} else {
  cat("ディレクトリは既に存在します: ", output_directory, "\n")
}

# 指定ディレクトリ内のすべてのJSONファイルをリストアップ
json_files <- list.files(json_directory, pattern = "\\.json$", full.names = TRUE)

# 新しいWord文書を作成
doc <- read_docx()

# 各JSONファイルの内容をループ処理で読み込み、すべての"text"要素をWord文書に追加
for (file_path in json_files) {
  # JSONファイルを読み込む
  json_data <- fromJSON(file_path)
  
  
  # "text"要素がリストやベクトル内に複数存在する場合、それらを反復処理する
  # JSONデータの構造に応じて、この部分を適切に調整する必要があるかもしれません
  if (is.list(json_data)) {
    # ファイル名を見出しとして追加
    doc <- doc %>% body_add_par(sprintf("日付: %s", basename(file_path)), style = "heading 2")
    
    text_content <- as.character()
    json_text <- unlist(json_data$text)
    if(length(json_text) > 0){
    for(i in 1:length(json_text)){
      text_content <- paste(text_content,json_text[i],sep="\n \n \n")
    }
            # "\n"でテキストを行に分割し、各行を新しい段落として追加
            lines <- strsplit(text_content, "\n")[[1]]
            for (line in lines) {
              doc <- doc %>% body_add_par(line, style = "Normal")
            }
          }
        }
  }



# Word文書をファイルに保存
print(doc, target = paste0(output_directory,"self-intro.docx"))

