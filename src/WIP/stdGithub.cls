VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdGithub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Based on: https://docs.github.com/en/rest
'Core docs: https://docs.github.com/en/free-pro-team@latest/rest/reference/repos#create-or-update-file-contents

'Other inspiration:
'https://github.com/github-tools/github/blob/master/lib/Repository.js
'


'Some small research:
'this.__fullname:  `${owner}/${repo}`
'Commit a repo:    this._request('POST', `/repos/${this.__fullname}/git/commits`, data, cb)
'Update repo:      this._request('PATCH', `/repos/${this.__fullname}`, options, cb);   (with  name, description, homepage, private status, has_issues, has_wiki, has_downloads, default_branch)
'Create branch:    this._request('POST', `/repos/${this.__fullname}/git/refs`, {sha: shaOfOldBranch, ref: `refs/heads/${newBranchName}` }}, cb);
'create pull req:  this._request('POST', `/repos/${this.__fullname}/pulls`, options, cb);
'update Pull Req:  this._request('PATCH', `/repos/${this.__fullname}/pulls/${number}`, options, cb);
'delete File:      this._request('DELETE', `/repos/${this.__fullname}/contents/${path}`, {message: "Some commit message", sha: shaOfBranchPath, branch:branchName}}, cb);
'move:             this.commit(oldSha, tree.sha, `Renamed '${oldPath}' to '${newPath}'`);   this.updateHead(`heads/${branch}`, commit.sha, true, cb));     [COMPLEX]
'create File:      this._request('PUT',`/repos/${this.__fullname}/contents/${path}, {branch: branchName, message: "some commit message", author: authorName, committer: committerName, content: "some content of file"})
'update file:      this._request('PUT',`/repos/${this.__fullname}/contents/${path}, {branch: branchName, message: "some commit message", author: authorName, committer: committerName, content: "some content of file"})
'





Public Function Create(ByVal sUsername as string, ByVal sToken as string, ByVal sRepository as string) as stdGithub

End Function
Public Function CreateFile() as boolean

End Function
Public Function UpdateFile() as boolean

End Function
Public Function DeleteFile() as boolean
    
End Function