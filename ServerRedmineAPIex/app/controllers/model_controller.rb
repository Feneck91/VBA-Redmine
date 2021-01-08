class ModelController < ApplicationController
  unloadable

  accept_api_auth :model
  respond_to :xml, :json
 
  def model
    logger.debug "Entering model"

    if (User.current.logged?)
      result = {}
      result['customField'] = CustomField.select('id, name, field_format, default_value, possible_values').where()
      result['status'] = IssueStatus.select('id, name')
      result['priority'] = Enumeration.select('id, name').where(:type => 'IssuePriority')
      result['users'] = User.select('id, login, firstname, lastname').where(:status => '1')
      result['groups'] = Group.select('id, lastname').where(:type => 'Group', :status => '1')
      respond_with(result, status: 200, location: nil)
    else
      respond_with({ 'Unauthorized' => 'Not allowed request, need to be logged on!' }, status: 401, location: nil)
    end
  end
end
