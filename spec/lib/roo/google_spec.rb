require 'spec_helper'

describe Roo::Google do
  let(:key) { '0AiokXJytm-hjdDhYbTNvZ3pDWm9oZm9yWURLX3ZoR2c' }

  describe '.new' do
    context 'given a username and password' do
      let(:user) { 'user' }
      let(:password) { 'password' }

      subject {
        Roo::Google.new(key, user: user, password: password)
      }

      it 'creates an instance' do
        VCR.use_cassette('google_drive') do
          expect(subject).to be_a(Roo::Google)
        end
      end
    end

    context 'given an access token' do
      let(:access_token) { 'ya29.AHES6ZR1kGjlmlLJG9skjpO0IjzQ6qDohXwFJclzD7mHI9xa-cFzlg' }

      subject {
        Roo::Google.new(key, access_token: access_token)
      }

      it 'creates an instance' do
        VCR.use_cassette('google_drive_access_token') do
          expect(subject).to be_a(Roo::Google)
        end
      end
    end
  end

  describe '.set' do
    let(:key) { '0AiVSev9-0p-PdEVrQldEY1I4dDJjVlpEQUkwblprcnc' }
    let(:user) { 'user' }
    let(:password) { "password" }

    subject {
      Roo::Google.new(key, user: user, password: password)
    }

    it 'records the value' do
      VCR.use_cassette('google_drive_set') do
        subject.cell(1, 1).should == '1x1'
        subject.cell(1, 2).should == '1x2'
        subject.cell(2, 1).should == '2x1'
        subject.cell(2, 2).should == '2x2'
        subject.set(1, 1, '1x1')
        subject.set(1, 2, '1x2')
        subject.set(2, 1, '2x1')
        subject.set(2, 2, '2x2')
        subject.cell(1, 1).should == '1x1'
        subject.cell(1, 2).should == '1x2'
        subject.cell(2, 1).should == '2x1'
        subject.cell(2, 2).should == '2x2'
      end 
    end
  end

end
