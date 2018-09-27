require 'podio'
require 'json'

config =  File.read('./config.json')
parsedJson = JSON.parse(config)

podioApiSecret = parsedJson['podioApiSecret ']
podioUser = parsedJson['podioUser']
podioPass = parsedJson['podioPass']
Podio.setup(:api_key => 'utsheets', :api_secret => podioApiSecret)
Podio.client.authenticate_with_credentials(podioUser, podioPass)

def startUp()
  begin
    responseBatchId = Podio.connection.post do |req|
    req.url '/item/app/1096164/export/xlsx'
    req.body = {:value => 'This is the text of the status message'}
    end

    batchId = responseBatchId.body['batch_id']
    sleep(180)
    
    responseFileId = Podio.connection.get("/batch/#{batchId}")

    fileId = responseFileId.body['file']['file_id']

    responseFileRaw = Podio.connection.get("/file/#{fileId}/raw")

    open('./files/PodioItemsExport.xlsx', 'w') { |f|
      f.puts responseFileRaw.body
    }
    puts('Podio items export written as Xlsx')
    puts('Info: fetchPodioItems.rb process successfully completed') 
    rescue StandardError => msg
    puts msg
    puts("Retrying in 120 seconds")
    sleep(120)
    startUp()
  end
end

startUp()