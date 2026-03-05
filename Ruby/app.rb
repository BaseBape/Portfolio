# frozen_string_literal: true

require "sinatra"
require "sinatra/reloader" if development?
require "json"
require "csv"

# HTMLフォルダ（同じ階層）を参照
html_root = File.expand_path("../HTML", __dir__)

configure do
  set :views, File.expand_path("views", html_root)
  set :public_folder, File.expand_path("public", html_root)
  set :inquiries_file, File.expand_path("inquiries.csv", __dir__)
end

helpers do
  def h(text)
    Rack::Utils.escape_html(text.to_s)
  end
end

get "/" do
  thanks = params[:thanks] == "1"
  flash = thanks ? { type: :success, messages: ["お問い合わせありがとうございます。内容を確認のうえ、ご連絡いたします。"] } : nil
  erb :index, locals: { flash: flash, form: {} }
end

post "/contact" do
  name    = params[:name]&.strip
  email   = params[:email]&.strip
  message = params[:message]&.strip

  errors = []
  errors << "お名前を入力してください" if name.to_s.empty?
  errors << "メールアドレスを入力してください" if email.to_s.empty?
  errors << "お問い合わせ内容を入力してください" if message.to_s.empty?

  if errors.any?
    status 422
    return erb :index, locals: {
      flash: { type: :error, messages: errors },
      form: { name: name, email: email, message: message }
    }
  end

  # CSVに保存（Rubyフォルダ内の inquiries.csv）
  File.open(settings.inquiries_file, "a") do |f|
    f.flock(File::LOCK_EX)
    f.puts(CSV.generate_line([Time.now.iso8601, name, email, message], force_quotes: true))
  end

  redirect to("/?thanks=1")
end
