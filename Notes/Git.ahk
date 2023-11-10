Notes_Git := [

	;; Github

	"gh repo fork",
	"gh repo fork [username]/[repository] --clone",

	;; Git

	"git stash", "
	(
		git stash
		git stash pop
		git stash save name
		git stash list
		git stash apply indexNumber
	)",

	"git change the name of a branch",
	"git branch -m oldName newName",

	"git change the name of the current branch",
	"git branch -M newName",

	"git restore code to the state of the remote repo", "
	(
		git fetch origin
		git reset --hard origin/main
	)",

	"git change your email",
	"git config --global user.email some_other_address",

	"git log a specific file -p",
	"git log -p file",

	"git new branch",
	"git checkout -b branch-name",

	"git merge when on the master branch",
	"git merge branch-name",

	"git squash merge",
	"git merge branch-name --squash",

	"git push a pull request",
	"git push upstream branch-name",

	"git how to not have to put in your credentials all the time",
	"git config --global credential.helper store",

	"git change the default main master branch name",
	"git config --global init.defaultBranch main",

	"git sync the credential manager",
	"git config --global credential.helper '/mnt/c/Program\ Files/Git/mingw64/libexec/git-core/git-credential-wincred.exe'",

	"git add remote to the current directory",
	"git remote add origin https://github.com/userName/repo.git",

	"git pick the upstream branch for future pushes and push to it",
	"git push -u origin main",

	"git undo only unstaged changes",
	"git restore --source=HEAD -- <file>",

	"git remove files locally but not on the remote",
	"git update-index --assume-unchanged files",

	"git interactive add", "
	(
		y: Stage the current hunk.
		n: Do not stage the current hunk and move to the next.
		q: Quit the interactive prompt and leave all remaining hunks unstaged.
		a: Stage the current hunk and all the following hunks in the current file.
		d: Skip the current hunk and all the following hunks in the current file.
		g: Select a hunk to go to.
		/: Search and navigate to a specific hunk, based on a regex pattern.
		j: Leave the current hunk unstaged and move to the next hunk.
		J: Leave the current hunk unstaged and move to the next hunk in the next file.
		k: Leave the current hunk unstaged and move to the previous hunk.
		K: Leave the current hunk unstaged and move to the previous hunk in the previous file.
		s: Split the current hunk into smaller hunks (if possible).
		e: Edit the current hunk manually.
	)",

	"git get short hash of a commit",
	"git rev-parse --short HEAD",

	"git get link to the remote repository",
	"git remote get-url origin",

	'git b_drive setup', '
	(
		1. Open Explorer and navigate on the b_drive to where you want to create a folder
		2. Create the folder:
			!!! Note the use of the '' => required if there are spaces in the folder names !!!
			['B:\San Francisco\Engineering\AutoHotkey\over_test']
		3. RClick new folder, open with Git Bash
		4. Initiate a global VIEW(able) by all repo; type the following:
			git init all
			[explanation: all (or world or everybody) - Same as group, but make the repository readable by all users]
		5. Navigate to the local repo folder in Git Bash; or open your local repo folder in Explorer, RClick, and open in Git Bash
			['C:\Users\bacona\OneDrive -['C:\Users\bacona\OneDrive - FM Global\Documents\AutoHotkey\Lib']
		6. Add a remote "shortcut"; type the following:
			!!! Note the added folder to the end of '\all' - since we created a folder that "all" can VIEW, and we initialized to that folder, so it's required !!!
			[git remote add b_over_test 'B:\San Francisco\Engineering\AutoHotkey\over_test\all']
		7. Push your local repo to the remote repo on the b_drive; type the following:
			[git push b_over_test]	
		Result:
				$ git push b_over_test
				Enumerating objects: 15708, done.
				Counting objects: 100% (15708/15708), done.
				Delta compression using up to 8 threads
				Compressing objects: 100% (8469/8469), done.
				Writing objects: 100% (15708/15708), 189.03 MiB | 435.00 KiB/s, done.
				Total 15708 (delta 6815), reused 15492 (delta 6730), pack-reused 0
				remote: Resolving deltas: 100% (6815/6815), done.
				remote: Checking connectivity: 15708, done.
				To B:\San Francisco\Engineering\AutoHotkey\over_test\all
				* [new branch]      main -> main
	)',
	"'git remote show <remote repo name>;' setup by [git remote add b_over_test ' 'B:\San Francisco\Engineering\AutoHotkey\over_test\all']",
	'git remote show b_over_test',
	''
]