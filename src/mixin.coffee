_ = require 'underscore'
_.mixin require('underscore.string').exports()

is_valid_date = (str)->
  regex_date = /(\d{4})-(\d{2})-(\d{2})/
  regex_date2 = /(\d{4})\/(\d{2})\/(\d{2})/
  regex_datetime = /(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/
  regex_datetime2 = /(\d{4})\/(\d{2})\/(\d{2}) (\d{2}):(\d{2})/

  if str.match(regex_date) or str.match(regex_date2) or str.match(regex_datetime) or str.match(regex_datetime2)
    return true
  else
    return false

is_number = (value)->
  if typeof(value) != 'number' && typeof(value) != 'string'
    false
  else
    value == parseFloat(value) && isFinite(value)

data_type = (value)->
  if is_number(value)
    return 'number'
  if is_valid_date(value)
    return 'date'
  if _.startsWith(value,'¥') or _.startsWith(value,'￥')
    return 'currency'
  if _.endsWith(value,'%')
    return 'percent'
  return 'other'

  
_.mixin is_valid_date:is_valid_date, is_number:is_number, data_type:data_type

