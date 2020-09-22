VERSION 5.00
Begin VB.Form frmQuest 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "?1"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   FillColor       =   &H0000FF00&
   FillStyle       =   0  'Solid
   Icon            =   "frmEllen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdD 
      Caption         =   "D"
      Height          =   495
      Left            =   975
      TabIndex        =   4
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   495
      Left            =   975
      TabIndex        =   3
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      Height          =   495
      Left            =   975
      TabIndex        =   2
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      Height          =   495
      Left            =   975
      TabIndex        =   1
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblLives 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1695
      TabIndex        =   8
      Top             =   4080
      Width           =   10005
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1695
      TabIndex        =   7
      Top             =   3480
      Width           =   10005
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1695
      TabIndex        =   6
      Top             =   2880
      Width           =   10005
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   10005
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numQuestions As Integer
Dim currentQuestion As Integer
Dim lives As Integer

Private Sub Question()

If currentQuestion = 1 Then
    frmQuest.Caption = "?1"
    lblQuestion.Caption = "You have just begun your quest. After traveling for 15 minutes, you see a haunted house. Do you:"
    lblA.Caption = "just stand there."
    lblB.Caption = "enter the haunted house."
    lblC.Caption = "scream."
    lblD.Caption = "get as far away from it as you can!"
End If

If currentQuestion = 2 Then
    frmQuest.Caption = "?2"
    lblQuestion.Caption = "Now you're in the Lost Desert. You had tripped on a Cheops Plant. Now you just realize that you are suffering from dehydration, hunger, and exhaustion. You come upon an inn. Do you:"
    lblA.Caption = "yell, 'HEY! GET ME SOMETHING TO EAT IN THERE!!!'"
    lblB.Caption = "go inside the inn."
    lblC.Caption = "just stay outside and lie down."
    lblD.Caption = "go for a walk."
End If

If currentQuestion = 3 Then
    frmQuest.Caption = "?3"
    lblQuestion.Caption = "Completely lost, you wander about. Eventually, you end up in Faerieland. Unfortunatly, you walk into Jhudora's Cave, and she's in a bad mood! Do you:"
    lblA.Caption = "scream after you've realized your mistake."
    lblB.Caption = "tell her how beautiful she is."
    lblC.Caption = "ask what's for dinner."
    lblD.Caption = "comment on her inventions."
End If

If currentQuestion = 4 Then
    frmQuest.Caption = "?4"
    lblQuestion.Caption = "After leaving Faerieland, you wander onto a ship unknowingly. Before you can realize what you've done, the ship has set sail. Do you:"
    lblA.Caption = "stay on the ship."
    lblB.Caption = "tell the captain to turn around."
    lblC.Caption = "smuggle yourself under some boxes."
    lblD.Caption = "sing sea shantees."
End If
If currentQuestion = 5 Then
    frmQuest.Caption = "?5"
    lblQuestion.Caption = "When you get off the ship onto Mystery Island, you run away from the pirates. Unfortunatly, you run into cannibals. They tie you up and then start boiling water over a fire. Do you:"
    lblA.Caption = "whistle."
    lblB.Caption = "cry."
    lblC.Caption = "laugh."
    lblD.Caption = "sing 'Old McCannibal had a Pot'."
End If
If currentQuestion = 6 Then
    frmQuest.Caption = "?6"
    lblQuestion.Caption = "Now you are on Terror Mountain. You decide to enter the Ice Caves. But you run into the sleeping Snowager. Do you:"
    lblA.Caption = "pet him."
    lblB.Caption = "try and steal some of his treasure."
    lblC.Caption = "hack him to pieces."
    lblD.Caption = "go buy Neggs."
End If
If currentQuestion = 7 Then
    frmQuest.Caption = "?7"
    lblQuestion.Caption = "You now are on Tyrannia. But you unknowingly go into the Lair of the Beast. Do you:"
    lblA.Caption = "snore loudly."
    lblB.Caption = "talk to the beast."
    lblC.Caption = "do the hokey pokey."
    lblD.Caption = "yodel."
End If
If currentQuestion = 8 Then
    frmQuest.Caption = "?8"
    lblQuestion.Caption = "You have made it to Meridell now. But you accidentally go to Darigan's place. You have interupted his plotting, and he throws you into the dungeon. Do you:"
    lblA.Caption = "sleep."
    lblB.Caption = "try to squeeze through the bars."
    lblC.Caption = "grab your cellphone and call a friend."
    lblD.Caption = "ask a nearby Turdle for help."
End If
If currentQuestion = 9 Then
    frmQuest.Caption = "?9"
    lblQuestion.Caption = "The Shoyru drops you off, and then flies away. You are hungry, and go into the Golden Dubloon. Do you buy:"
    lblA.Caption = "a red apple."
    lblB.Caption = "chicken."
    lblC.Caption = "a pyramicake."
    lblD.Caption = "mince pie."
End If
If currentQuestion = 10 Then
    frmQuest.Caption = "?10"
    lblQuestion.Caption = "You drop deep into Maraqua. Unfortunatly, a sea creature grabs you and intends to eat you. Do you:"
    lblA.Caption = "use the knife that you saw."
    lblB.Caption = "attack."
    lblC.Caption = "tell him to let go."
    lblD.Caption = "play dead."
End If
If currentQuestion = 11 Then
    frmQuest.Caption = "?11"
    lblQuestion.Caption = "You have swum to Neopia Central. You then find your way to your house. You are very happy! What do you do next?"
    lblA.Caption = "watch T.V."
    lblB.Caption = "take a nap."
    lblC.Caption = "go out to eat."
    lblD.Caption = "watch the sunset."
End If
End Sub

Private Sub cmdA_Click()

If currentQuestion = 10 Then
    If MsgBox("Are you mad?!!! YOU NEVER SAW A KNIFE! IF YOU DID I WOULD'VE TOLD YOU NOW, WOULDN'T I HAVE?!!!", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 9 Then
    If MsgBox("Unfortunatly, it is poisinous.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 8 Then
    If MsgBox("You are rudely awakened by a guard. He summons you to the execution grounds...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 7 Then
    If MsgBox("You startle the beast, and it attacks you with its sharp talons...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 6 Then
    If MsgBox("The Snowager wakes up and bites you with its poisonous fangs...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 5 Then
    If MsgBox("Some Tree Cybunnies hear you, and they scare the cannibals away. Using Eyries, their transports, they fly you to Terror Mountain, then leave.", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 4 Then
    If MsgBox("Soon some crew members see you. It turns out they're pirates, and they cut your throat.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 3 Then
    If MsgBox("Jhudora immediatly sees you. Using her powers, she lulls you into an eternal sleep...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 2 Then
    If MsgBox("An armed Techo comes out. Thinking you're a murderer, he attacks...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 1 Then
    If MsgBox("After staring at the house too long, you become dazed. You then become unconscious. You never wake up...", vbOKOnly) = vbOK Then lives = lives - 1
End If

If lives = 0 Then End
If currentQuestion = numQuestions + 1 Then End

lblLives.Caption = lives
Question

End Sub

Private Sub cmdB_Click()

If currentQuestion = 10 Then
    If MsgBox("You struggle, but the creature is too strong. He then eats you...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 9 Then
    If MsgBox("Unfortunatly, it is poisonous.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 8 Then
    If MsgBox("You try to squeeze through the bars, but a guard sees you and soon ends it forever...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 7 Then
    If MsgBox("You start a pleasant conversation with the beast. He explains he attacks others to protect his home. Soon, you say you must be on your way, and leave.", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 6 Then
    If MsgBox("Unfortunatly, the Snowager hears you rumaging through his treasure, and squeezes you to death...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 5 Then
    If MsgBox("The cannibals just continue boiling the water. Soon, it is ready, and they put you in...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 4 Then
    If MsgBox("He only laughs, and you see his eye patch and wooden leg. You look up, and see a pirate flag. Some crew members tie you up, and drop you into the sea...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 3 Then
    If MsgBox("Jhudora screams loudly. 'I'LL KILL YOU, YOU MORTAL FOOL!' But her noise has brought in her sisters, who soon lead you to safety. Soon, you are on your way again.", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 2 Then
    If MsgBox("You make reseravations and then sit down for lunch. You are served some delicious juice and Sphinx Links. After a good rest, you continue your journey...", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 1 Then
    If MsgBox("After you enter, you look up. You see an axe falling towards you...", vbOKOnly) = vbOK Then lives = lives - 1
End If

If lives = 0 Then End
If currentQuestion = numQuestions + 1 Then End

lblLives.Caption = lives
Question

End Sub

Private Sub cmdC_Click()

If currentQuestion = 10 Then
    If MsgBox("The creature lets go. You float to the surface, and start swimming...", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 9 Then
    If MsgBox("You eat, and then go outside. Suddenly, an evil Pawkeet picks you up and drops you into the water...", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 8 Then
    If MsgBox("You forgot that you don't have a cellphone, and accidentally pick up a knife instead. Unfortunatly, it slips and goes right towards your heart...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 7 Then
    If MsgBox("You have wasted too much time, and the Jelly Chia enters the caves...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 6 Then
    If MsgBox("You yell, which wakes up the Snowager, and are about to hack him to bits, when you realize you don't have a weapon...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 5 Then
    If MsgBox("The cannibals look at you strangly, then continue boiling. Soon, the water is ready...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 4 Then
    If MsgBox("It is dusty under the boxes. You sneeze, which attracts some crew members. They find you, and feed you to their pet shark.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 3 Then
    If MsgBox("Jhudora laughs wickedly. 'You are!'", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 2 Then
    If MsgBox("You fall asleep. When you wake up, you're choking from dehydration. Your stomach is practically screaming. Then, your vision seems to fade away...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 1 Then
    If MsgBox("You scream. But only the Jelly Chia hears you.", vbOKOnly) = vbOK Then lives = lives - 1
End If

If lives = 0 Then End
If currentQuestion = numQuestions + 1 Then End

lblLives.Caption = lives
Question

End Sub
Private Sub cmdD_Click()

If currentQuestion = 10 Then
    If MsgBox("The creature eats you anyway.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 9 Then
    If MsgBox("Unfortunatly, it is poisonous.", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 8 Then
    If MsgBox("The Turdle is very understanding, and grinds through the bars. A Shoyru then flies you to Krawk Island.", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 7 Then
    If MsgBox("A very unfortunate combination. Both the beast and Jelly Chia are after you...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 6 Then
    If MsgBox("You go to the Negg Shop, and find the Negg Faerie about to go to lunch. She invites you, and you accept. Lunch is at Tyrannia...", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 5 Then
    If MsgBox("For a bit the cannibals dance. But soon they get back to work. Soon you will be their dinner...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 4 Then
    If MsgBox("Luckily, you sing the pirates' favorite song. They think you're one of the crew, and you sail all the way to Mystery Island.", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If
If currentQuestion = 3 Then
    If MsgBox("She says 'My inventions may be grand, but I take an instant disliking to you!' She then opens a bottle of Noxious Nectar to your face. 'This is my latest!' You look away, right before she throws her sword at you...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 2 Then
    If MsgBox("You go for a simple walk. But you run into the hungry Jelly Chia...", vbOKOnly) = vbOK Then lives = lives - 1
End If
If currentQuestion = 1 Then
    If MsgBox("You run away from the haunted house. You turn around, but you trip on something...", vbOKOnly) = vbOK Then currentQuestion = currentQuestion + 1
End If

If lives = 0 Then End
If currentQuestion = numQuestions + 1 Then End

lblLives.Caption = lives
Question

End Sub

Private Sub Form_Load()

numQuestions = 10
currentQuestion = 1
lives = 3

If MsgBox("You are lost in Neopia and need to find your way back home to Neopia Central. Unfortunatly, the Jelly Chia is chasing you the whole way there. He is a Chia made of jelly who eats anything in his way. Will you escape him? Your journey starts in the Haunted Woods...", vbOKOnly) = vbOK Then
End If

Question

End Sub
