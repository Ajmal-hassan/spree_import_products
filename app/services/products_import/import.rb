require 'roo'
require 'open-uri'
require 'addressable/uri'
class ProductNotCreatedError < ::StandardError
end

class ProductsImport::Import
  attr_accessor :imported_sheet, :spree_current_user

  def initialize(file_path, spree_current_user, import_log, current_store)
    @import_log = import_log
    imported_file = Roo::Spreadsheet.open(file_path, extension: :xlsx)
    @imported_sheet = imported_file.sheet(0)
    @spree_current_user = spree_current_user
    @current_store = current_store
    @taxonomy = Spree::Taxonomy.where(name: 'Categories', store: @current_store ).first_or_create!
    @country =  Spree::Country.find_by(iso: 'US')
  end

  def perform_import
    @failed = []
    @product_name = nil
    # @import_log.update_columns(state: 'running')
    header = imported_sheet.row(1)
    (2..imported_sheet.last_row).each do |index|
      row = [header, imported_sheet.row(index)].transpose.to_h
      create_product(row, index) unless row.values.uniq.first == nil
    end
    # @import_log.update_column(:state, 'finished')
    # create_and_send_product_import_report
  end

  private

  def create_product(row, index)
    shipping_category = find_or_create_shipping_category(row['ShippingCategory']) if row['ShippingCategory']
    tax_category = find_or_create_tax_category(row['TaxCategory']) if row['TaxCategory']
    # Product find from sku for updating the variant if present
    if row['Name'].present?
      @product_name = row['Name']
      @product = Spree::Product.find_or_initialize_by(name: row['Name'])
      @product.price = row['Price']
      @product.shipping_category = shipping_category
      @product.description = row['Description']
      @product.available_on = Time.parse(row['AvailableOn'].to_s)
      @product.tax_category = tax_category
      @product.sku = row['SKU']  if row['SKU']
      @product.cost_price = row['Price']
      @product.master.currency = @current_store.default_currency
      @product.master.cost_currency = @current_store.default_currency
      if @product.new_record?
        @product.stores << @current_store
        @product.status = 'active'
        if @product.save!
          # save_properties(row) if row['ProductProperties']
          save_taxons(row, @product) if row['Taxons']
          if row['option_type'].present?
            save_variants(@product, row)
          end
          if row['image'].present?
            save_product_img(@product.master, row['image'])
          end
          if row['option_type'].nil? && row['additional_image'].present?
            save_product_img(@product.master, row['additional_image'])
          end

        end
      elsif @product.persisted?
        @product.status = 'active'
        save_variants(@product, row)
        save_taxons(row, @product) if row['Taxons']
        if row['image'].present?
          save_product_img(@product.master, row['image'])
        end
        if row['option_type'].nil? && row['additional_image'].present?
          save_product_img(@product.master, row['additional_image'])
        end
      end
    elsif row['name'].nil? && row['model'].present?
      @product = Spree::Product.find_by(name: @product_name)
      if @product.present?
        if row['option_type'].present?
          save_variants(@product, row)
        end
        if row['Taxons'].present?
          save_taxons(row, @product) if row['Taxons']
        end
        if row['option_type'].nil? && row['additional_image'].present?
          save_product_img(@product.master, row['additional_image'])
        end
      end

    end
  rescue ActiveRecord::RecordInvalid => e
    handle_exception(row, e, index)
  rescue ProductNotCreatedError => e
    handle_exception(row, e, index)
  rescue Exception => e
    handle_exception(row, e, index)
  end

  def error_details(row,e, index)
    {"row_#{index}": { product_name: row['Name'], message: e.message }}
  end

  def handle_exception(product, e, index)
    @failed.push(product['Name'])
    data = error_details(product,e,index)
    @import_log.error_details.merge!(data)
    @import_log.increment!(:error_row_count)
    @import_log.save
  end

  def save_taxons(row, product)
    categories_taxon = @taxonomy.taxons.where(name: I18n.t('spree.taxonomy_categories_name')).first_or_create!

    if row['Taxons'].split('///').size == 1
      parent = categories_taxon.children.where(name: row['Taxons']).first_or_create!
      product.taxons << parent unless product.taxons.find_by(name: parent.name)
    else
      final_taxons = nil
      row['Taxons'].split('///').each do |t|
        categories_taxon = categories_taxon.children.where(name: t).first_or_create!
        final_taxons = categories_taxon
      end
      product.taxons << final_taxons unless product.taxons.find_by(name: final_taxons.name)
    end
    # end
    product.save
  end

  def save_variants(product, row) #colour:blue;size:small | colour:red;size:medium | colour:green;size:small
    if row['option_type'].present? && row['option_value'].present?
      option_type = Spree::OptionType.find_or_create_by(name: row['option_type'], presentation: row['option_type'].titleize)
      option_value = Spree::OptionValue.find_or_create_by(name: row['option_value'], presentation: row['option_value'].titleize, option_type: option_type)
      if option_type.present?
        if product.option_types.present?
          existing_option_type = product.option_types.find(option_type.id)
          if existing_option_type.nil?
            product.option_types << option_type
          end
        else
          product.option_types << option_type
        end
        already_present_variant = product.variants.find_by(sku: option_value.name) || nil
        quantity = row['option_quantity']
        image_url = row['additional_image']
        if !already_present_variant.present?
          product_variant = product.variants.new(cost_price: product.cost_price, sku: option_value.name, cost_currency: @current_store.default_currency, currency: @current_store.default_currency, track_inventory: false)
          product_variant.option_values << option_value
          if product_variant.save!
            update_stock_on_hand_and_track_inventory(product_variant, quantity)
            if image_url.present?
              add_images(product_variant, image_url)
            end
          end
        else
          already_present_variant.update(cost_price: product.cost_price, price: product.price)
          update_stock_on_hand_and_track_inventory(already_present_variant, quantity)
          if image_url.present?
            add_images(already_present_variant,image_url)
          end

        end

      end
    end

    # variant = product.variants.find_or_initialize_by(option_values: [option_value])
    # variant.price = row['price'].to_f
    # variant.stock_items.find_or_initialize_by(stock_location: default_stock_location).update(count_on_hand: row['option_quantity'].to_i)
    # attach_image(variant, row['image'])

    # variant.save!
    # row['Variants'].split('|').each_with_index do |options, index| #["colour:blue;size:small ", " colour:red;size:medium ", " colour:green;size:small"]
    #   #{"colour"=>"blue", "size"=>"small "}
    #   option_values = []
    #   # variant_hash = Hash[variant.split(';').map { |x| [x.split(':').first, x.split(':').second] }]
    #   options.split(';').each do |v| #["colour:blue", "size:small "]
    #     option_type_and_value = v.strip.split(':')
    #     if Spree::Store._reflect_on_association(:option_types)
    #       option_type = @current_store.option_types.find_or_create_by!(name: option_type_and_value.first, presentation: option_type_and_value.first.capitalize)
    #       option_values =  option_values + [option_type.option_values.find_or_create_by!( name: option_type_and_value.second,presentation: option_type_and_value.second.capitalize,)]
    #     else
    #       option_type = Spree::OptionType.find_or_create_by(name:option_type_and_value.first, presentation:option_type_and_value.first.capitalize)
    #       option_values = option_values + [option_type.option_values.find_or_create_by(name:option_type_and_value.second, presentation:option_type_and_value.second.capitalize)]
    #     end
    #   end
    #
    #   # Variant find from sku for updating the variant if present

    # end
  end

  def add_images(variant, image_url)
    save_product_img(variant, image_url)
  end

  def find_or_create_shipping_category(name)
    if Spree::Store._reflect_on_association(:shipping_categories)
      @current_store.shipping_categories.find_or_create_by(name: name + "_#{@current_store.id}")
    else
      Spree::ShippingCategory.find_or_create_by(name: name + "_#{@current_store.id}")
    end
  end

  def find_or_create_tax_category(name)
    if Spree::Store._reflect_on_association(:tax_categories)
      @current_store.tax_categories.find_or_create_by(name: name + "_#{@current_store.id}")
    else
      Spree::TaxCategory.find_or_create_by(name: name + "_#{@current_store.id}")
    end
  end

  def save_properties(product)
    properties = product['ProductProperties'].split('|')
    properties.each do |p|
      final_property = p.split(':')
      if Spree::Store._reflect_on_association(:properties)
        property = @current_store.properties.find_or_create_by!(name: final_property[0], presentation: final_property[0])
        product_property = Spree::ProductProperty.where(product: @product, property: property).first_or_initialize
        product_property.value = final_property[1]
        product_property.save!
      else
        property = Spree::Property.find_or_create_by(name:final_property[0], presentation:final_property[0])
        product_property = Spree::ProductProperty.where(product: @product, property:property).first_or_initialize
        product_property.value = final_property[1]
        product_property.save
      end
    end
  end

  def update_stock_on_hand_and_track_inventory(variant, quantity)
    stock_location = Spree::StockLocation.where(default: true).first
    unless stock_location.present?
      stock_location = Spree::StockLocation.where(name: 'Default', default: true, address1: 'Example Street',
                                                  city: 'City', zipcode: '12345', country: @country, state: @country.states.first,
                                                  active: true, propagate_all_variants: true).first_or_create!
    end
    stock_movement = stock_location.stock_movements.build(quantity: quantity.to_i)
    stock_movement.stock_item = stock_location.set_up_stock_item(variant)
    stock_movement.stock_item.update(count_on_hand: quantity.to_i)
    stock_movement.save
    variant.update(track_inventory: false)
    stock_movement.stock_item.update(backorderable: true)
  end

  def save_product_img(variant, img_url)
    ActiveRecord::Base.transaction do
      begin
        filename = img_url.split('/').last
        existing_img = variant.images.select { |img| filename == img.attachment.blob.filename.to_s }
        if !existing_img.present?
          new_img_url = build_full_image_url(img_url)
          image = URI.parse(new_img_url).open
          variant.images.create!(attachment: { io: image, filename: filename })
        end
      rescue Exception => e
        puts "Exception #{e} for #{image} and product - #{variant.product.id}"
      end
    end
  end

  def build_full_image_url(relative_path, base_url = "https://www.tanstartrade.com/image/cache", resolution = "-1500x1500")
    normalized_path = relative_path.gsub(/^\//, '').gsub(/\/+/, '/')
    transformed_path = normalized_path.gsub(/(\.\w+)$/, "#{resolution}\\1")
    full_url = "#{base_url}/#{transformed_path}"
    Addressable::URI.parse(full_url).normalize.to_s
  end
end